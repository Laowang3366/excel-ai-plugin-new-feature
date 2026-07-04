/**
 * RAG 知识增强层 — 单元测试
 *
 * 覆盖：EmbeddingService、SqliteStore、DocumentParser、
 *       TextChunker、KnowledgeIndexer、Retriever
 *
 * SqliteStore 使用 :memory: 数据库，无需真实文件。
 * EmbeddingService 使用 mock fetch。
 */

import { describe, it, expect, beforeEach, afterEach, vi } from "vitest";
import { randomUUID } from "crypto";
import * as path from "path";
import * as fs from "fs";
import * as os from "os";
import JSZip from "jszip";

// ============================================================
// 类型和模块导入
// ============================================================

import type { KnowledgeEntry, KnowledgeSource, KnowledgeQuery, KnowledgeResult } from "./types";
import { openSqliteDatabase } from "../storage/nodeSqlite";

// ============================================================
// EmbeddingService 测试
// ============================================================

describe("EmbeddingService", () => {
  let EmbeddingService: typeof import("./embeddingService").EmbeddingService;

  beforeEach(async () => {
    const mod = await import("./embeddingService");
    EmbeddingService = mod.EmbeddingService;
  });

  it("should create with default model config for provider", () => {
    const svc = new EmbeddingService({
      provider: "openai",
      baseUrl: "https://api.openai.com",
      apiKey: "sk-test",
    });
    expect(svc.getModelName()).toBe("text-embedding-3-small");
    expect(svc.getDimensions()).toBe(1536);
  });

  it("should fallback to custom model when provided", () => {
    const svc = new EmbeddingService({
      provider: "openai",
      baseUrl: "https://api.openai.com",
      apiKey: "sk-test",
      model: "custom-embedding-model",
    });
    expect(svc.getModelName()).toBe("custom-embedding-model");
  });

  it("should throw on empty text", async () => {
    const svc = new EmbeddingService({
      provider: "openai",
      baseUrl: "https://api.openai.com",
      apiKey: "sk-test",
    });
    await expect(svc.embed("")).rejects.toThrow("不能为空");
    await expect(svc.embed("   ")).rejects.toThrow("不能为空");
  });

  it("should cache embeddings for identical text", async () => {
    // Mock fetch to count calls
    let fetchCount = 0;
    const originalFetch = globalThis.fetch;
    globalThis.fetch = async (url: any, init?: any) => {
      fetchCount++;
      return new Response(JSON.stringify({
        data: [{ embedding: [0.1, 0.2, 0.3] }],
        model: "text-embedding-3-small",
        usage: { prompt_tokens: 4, total_tokens: 4 },
      }), { status: 200, headers: { "Content-Type": "application/json" } });
    };

    try {
      const svc = new EmbeddingService({
        provider: "openai",
        baseUrl: "https://api.openai.com",
        apiKey: "sk-test",
      });

      const r1 = await svc.embed("测试文本");
      const r2 = await svc.embed("测试文本"); // 应命中缓存
      const r3 = await svc.embed("不同的文本"); // 应再次调用 API

      expect(r1).toEqual(r2); // 缓存命中，结果相同
      expect(fetchCount).toBe(2); // 只调用了 2 次 API
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it("should normalize provider base URL for embedding requests", async () => {
    const requestedUrls: string[] = [];
    const originalFetch = globalThis.fetch;
    globalThis.fetch = async (url: any) => {
      requestedUrls.push(String(url));
      return new Response(JSON.stringify({
        data: [{ embedding: [0.1, 0.2, 0.3] }],
      }), { status: 200, headers: { "Content-Type": "application/json" } });
    };

    try {
      const rootUrlService = new EmbeddingService({
        provider: "openai",
        baseUrl: "https://api.example.com",
        apiKey: "sk-test",
      });
      await rootUrlService.embed("根地址");

      const chatUrlService = new EmbeddingService({
        provider: "openai",
        baseUrl: "https://api.example.com/v1/chat/completions",
        apiKey: "sk-test",
      });
      await chatUrlService.embed("聊天接口地址");

      expect(requestedUrls).toEqual([
        "https://api.example.com/v1/embeddings",
        "https://api.example.com/v1/embeddings",
      ]);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it("should summarize HTML embedding errors without leaking page markup", async () => {
    const originalFetch = globalThis.fetch;
    globalThis.fetch = async () => {
      return new Response(
        `<html><body><div data-component="top">404</div><script>window.__NEXT_DATA__={}</script></body></html>`,
        { status: 404, headers: { "Content-Type": "text/html" } }
      );
    };

    try {
      const svc = new EmbeddingService({
        provider: "openai",
        baseUrl: "https://api.example.com",
        apiKey: "sk-test",
      });

      await expect(svc.embed("测试")).rejects.toThrow("服务返回了网页内容");
      await expect(svc.embed("测试")).rejects.not.toThrow("data-component");
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it("should handle batch embedding with caching", async () => {
    let fetchCount = 0;
    const originalFetch = globalThis.fetch;
    globalThis.fetch = async () => {
      fetchCount++;
      return new Response(JSON.stringify({
        data: [
          { embedding: [0.1, 0.2, 0.3] },
          { embedding: [0.4, 0.5, 0.6] },
        ],
        model: "text-embedding-3-small",
        usage: { prompt_tokens: 8, total_tokens: 8 },
      }), { status: 200 });
    };

    try {
      const svc = new EmbeddingService({
        provider: "openai",
        baseUrl: "https://api.openai.com",
        apiKey: "sk-test",
      });

      // 先缓存一个
      await svc.embed("文本A");

      // 批处理：一个缓存命中，一个需要调用 API
      const results = await svc.embedBatch(["文本A", "文本B"]);

      expect(results).toHaveLength(2);
      expect(fetchCount).toBe(2); // embed 调用一次, embedBatch 调用一次（只处理未缓存的）
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it("should clear cache on request", async () => {
    let fetchCount = 0;
    const originalFetch = globalThis.fetch;
    globalThis.fetch = async () => {
      fetchCount++;
      return new Response(JSON.stringify({
        data: [{ embedding: [0.1, 0.2, 0.3] }],
        model: "text-embedding-3-small",
      }), { status: 200 });
    };

    try {
      const svc = new EmbeddingService({
        provider: "openai",
        baseUrl: "https://api.openai.com",
        apiKey: "sk-test",
      });

      await svc.embed("测试");
      svc.clearCache();
      await svc.embed("测试"); // 清空缓存后应再次调用

      expect(fetchCount).toBe(2);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

// ============================================================
// SqliteStore 测试（使用 :memory: 数据库）
// ============================================================

describe("SqliteStore", () => {
  let SqliteStore: typeof import("./sqliteStore").SqliteStore;
  let store: any;

  const sampleEntry = (overrides?: Partial<KnowledgeEntry>): KnowledgeEntry => ({
    id: randomUUID(),
    source: "document",
    sourcePath: "/test/file.xlsx",
    sourceName: "file.xlsx",
    sourceType: "xlsx",
    chunkIndex: 0,
    content: "测试内容",
    metadata: { sheetName: "Sheet1", rowCount: 10 },
    embedding: [0.1, 0.2, 0.3],
    indexedAt: Date.now(),
    tokenCount: 10,
    ...overrides,
  });

  beforeEach(async () => {
    const mod = await import("./sqliteStore");
    SqliteStore = mod.SqliteStore;
    store = new SqliteStore(":memory:");
    await store.init();
  });

  afterEach(() => {
    store.close();
  });

  it("should initialize tables on construction", () => {
    // 建表是幂等的，构造时已执行
    const count = store.countEntries();
    expect(count).toBe(0);
  });

  it("should migrate legacy entry schema before creating embedding profile index", async () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), `knowledge-legacy-${Date.now()}-`));
    const dbPath = path.join(tmpDir, "knowledge.db");
    let legacyStore: InstanceType<typeof SqliteStore> | undefined;

    try {
      const legacyDb = openSqliteDatabase(dbPath);
      try {
        legacyDb.exec(`
          CREATE TABLE knowledge_entries (
            id          TEXT PRIMARY KEY,
            source      TEXT NOT NULL,
            source_path TEXT NOT NULL,
            source_name TEXT NOT NULL,
            source_type TEXT NOT NULL,
            chunk_index INTEGER DEFAULT 0,
            content     TEXT NOT NULL,
            metadata    TEXT DEFAULT '{}',
            embedding   TEXT,
            indexed_at  INTEGER NOT NULL,
            token_count INTEGER DEFAULT 0
          );

          CREATE TABLE knowledge_sources (
            source_path   TEXT PRIMARY KEY,
            source_name   TEXT NOT NULL,
            source_type   TEXT NOT NULL,
            entry_count   INTEGER DEFAULT 0,
            first_indexed INTEGER NOT NULL,
            last_indexed  INTEGER NOT NULL,
            file_hash     TEXT DEFAULT ''
          );

          INSERT INTO knowledge_entries
            (id, source, source_path, source_name, source_type,
             chunk_index, content, metadata, embedding, indexed_at, token_count)
          VALUES
            ('legacy-entry', 'document', '/legacy/file.txt', 'file.txt', 'txt',
             0, 'legacy content', '{}', '[1,0,0]', 1000, 2);

          INSERT INTO knowledge_sources
            (source_path, source_name, source_type, entry_count,
             first_indexed, last_indexed, file_hash)
          VALUES
            ('/legacy/file.txt', 'file.txt', 'txt', 1, 1000, 1000, 'hash');
        `);
      } finally {
        legacyDb.close();
      }

      legacyStore = new SqliteStore(dbPath);
      await expect(legacyStore.init()).resolves.toBeUndefined();
      expect(legacyStore.countEntries()).toBe(1);
      expect(legacyStore.listSources()).toHaveLength(1);
      expect(legacyStore.getEntry("legacy-entry")?.embedding).toEqual([1, 0, 0]);

      legacyStore.close();
      legacyStore = undefined;

      const migratedDb = openSqliteDatabase(dbPath);
      try {
        const columns = migratedDb
          .prepare("PRAGMA table_info(knowledge_entries)")
          .all()
          .map((row: any) => row.name);
        const indexes = migratedDb
          .prepare("PRAGMA index_list(knowledge_entries)")
          .all()
          .map((row: any) => row.name);

        expect(columns).toContain("embedding_provider");
        expect(columns).toContain("embedding_model");
        expect(columns).toContain("embedding_dimensions");
        expect(indexes).toContain("idx_entries_embedding_profile");
      } finally {
        migratedDb.close();
      }
    } finally {
      if (legacyStore?.isInitialized()) {
        legacyStore.close();
      }
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });

  it("should insert and retrieve a single entry", () => {
    const entry = sampleEntry();
    store.insertEntry(entry);

    const retrieved = store.getEntry(entry.id);
    expect(retrieved).not.toBeNull();
    expect(retrieved!.id).toBe(entry.id);
    expect(retrieved!.content).toBe(entry.content);
    expect(retrieved!.metadata.sheetName).toBe("Sheet1");
    expect(retrieved!.embedding).toEqual([0.1, 0.2, 0.3]);
  });

  it("should bulk insert entries with transaction", () => {
    const entries = Array.from({ length: 100 }, (_, i) =>
      sampleEntry({
        id: randomUUID(),
        chunkIndex: i,
        content: `条目 ${i}`,
      })
    );

    store.bulkInsert(entries);
    expect(store.countEntries()).toBe(100);
  });

  it("should delete a single entry", () => {
    const entry = sampleEntry();
    store.insertEntry(entry);
    expect(store.countEntries()).toBe(1);

    store.deleteEntry(entry.id);
    expect(store.getEntry(entry.id)).toBeNull();
    expect(store.countEntries()).toBe(0);
  });

  it("should delete all entries by source path", () => {
    store.bulkInsert([
      sampleEntry({ id: randomUUID(), sourcePath: "/test/a.xlsx" }),
      sampleEntry({ id: randomUUID(), sourcePath: "/test/a.xlsx" }),
      sampleEntry({ id: randomUUID(), sourcePath: "/test/b.xlsx" }),
    ]);
    expect(store.countEntries()).toBe(3);

    store.deleteSource("/test/a.xlsx");
    expect(store.countEntries()).toBe(1);
  });

  it("should return correct cosine similarity ranking", () => {
    // 插入三条不同向量
    store.bulkInsert([
      sampleEntry({ id: "id-1", content: "A", embedding: [1, 0, 0] }),
      sampleEntry({ id: "id-2", content: "B", embedding: [0, 1, 0] }),
      sampleEntry({ id: "id-3", content: "C", embedding: [0.9, 0.1, 0] }),
    ]);

    // 搜索接近 [1, 0, 0] 的向量
    // 注意：entry-3 的 embedding [0,1,0] dot [1,0,0] = 0，会被 score>0 过滤
    const results = store.searchByVector([1, 0, 0], 3);

    expect(results).toHaveLength(2);
    // A 应该是最相关的
    expect(results[0].entry.id).toBe("id-1");
    expect(results[0].score).toBeCloseTo(1, 2);
    // C 第二
    expect(results[1].entry.id).toBe("id-3");
    expect(results[1].score).toBeGreaterThan(0.9);
  });

  it("should filter vector search by source", () => {
    store.bulkInsert([
      sampleEntry({ id: "id-1", source: "workbook", sourcePath: "/test/a.xlsx", content: "A", embedding: [1, 0, 0] }),
      sampleEntry({ id: "id-2", source: "document", sourcePath: "/test/b.txt", content: "B", embedding: [1, 0, 0] }),
    ]);

    const results = store.searchByVector([1, 0, 0], 10, { sourceFilter: ["workbook"] });
    expect(results).toHaveLength(1);
    expect(results[0].entry.id).toBe("id-1");
  });

  it("should ignore vectors with mismatched dimensions", () => {
    store.bulkInsert([
      sampleEntry({ id: "id-1", content: "A", embedding: [1, 0, 0] }),
      sampleEntry({ id: "id-2", content: "B", embedding: [1, 0] }),
    ]);

    const results = store.searchByVector([1, 0, 0], 10);
    expect(results.map((result: KnowledgeResult) => result.entry.id)).toEqual(["id-1"]);
  });

  it("should filter vector search by embedding profile", () => {
    store.bulkInsert([
      sampleEntry({
        id: "id-1",
        content: "current",
        embedding: [1, 0, 0],
        embeddingProvider: "openai",
        embeddingModel: "text-embedding-3-small",
        embeddingDimensions: 3,
      }),
      sampleEntry({
        id: "id-2",
        content: "old",
        embedding: [1, 0, 0],
        embeddingProvider: "qwen",
        embeddingModel: "text-embedding-v2",
        embeddingDimensions: 3,
      }),
    ]);

    const results = store.searchByVector([1, 0, 0], 10, {
      embeddingProfile: {
        provider: "openai",
        model: "text-embedding-3-small",
        dimensions: 3,
      },
    });

    expect(results.map((result: KnowledgeResult) => result.entry.id)).toEqual(["id-1"]);
  });

  it("should search by keyword using LIKE", () => {
    store.bulkInsert([
      sampleEntry({ id: "id-1", content: "销售数据报表" }),
      sampleEntry({ id: "id-2", content: "财务报表分析" }),
      sampleEntry({ id: "id-3", content: "项目进度跟踪" }),
    ]);

    const results = store.searchByKeyword(["销售"], 10);
    expect(results).toHaveLength(1);
    expect(results[0].content).toContain("销售");
  });

  it("should manage sources", () => {
    const source: KnowledgeSource = {
      sourcePath: "/test/file.xlsx",
      sourceName: "file.xlsx",
      sourceType: "xlsx",
      entryCount: 5,
      firstIndexed: 1000,
      lastIndexed: 2000,
      fileHash: "abc123",
    };

    store.upsertSource(source);

    const listed = store.listSources();
    expect(listed).toHaveLength(1);
    expect(listed[0].sourcePath).toBe("/test/file.xlsx");

    const retrieved = store.getSource("/test/file.xlsx");
    expect(retrieved).not.toBeNull();
    expect(retrieved!.fileHash).toBe("abc123");
  });

  it("should show indexed entries as sources even when source summaries are missing", () => {
    store.bulkInsert([
      sampleEntry({
        id: "summary-id-1",
        sourcePath: "/test/orphan.xlsx",
        sourceName: "orphan.xlsx",
        sourceType: "xlsx",
        indexedAt: 1000,
      }),
      sampleEntry({
        id: "summary-id-2",
        sourcePath: "/test/orphan.xlsx",
        sourceName: "orphan.xlsx",
        sourceType: "xlsx",
        indexedAt: 2000,
      }),
    ]);

    const listed = store.listSources();

    expect(listed).toHaveLength(1);
    expect(listed[0]).toMatchObject({
      sourcePath: "/test/orphan.xlsx",
      sourceName: "orphan.xlsx",
      sourceType: "xlsx",
      entryCount: 2,
      firstIndexed: 1000,
      lastIndexed: 2000,
    });
  });

  it("should handle empty vector search gracefully", () => {
    const results = store.searchByVector([1, 0, 0], 10);
    expect(results).toHaveLength(0);
  });
});

// ============================================================
// DocumentParser 测试
// ============================================================

describe("DocumentParser", () => {
  let DocumentParser: typeof import("./documentParser").DocumentParser;
  let parser: any;

  beforeEach(async () => {
    const mod = await import("./documentParser");
    DocumentParser = mod.DocumentParser;
    parser = new DocumentParser();
  });

  it("should detect supported file types", () => {
    expect(parser.isSupported("test.xlsx")).toBe(true);
    expect(parser.isSupported("test.csv")).toBe(true);
    expect(parser.isSupported("test.md")).toBe(true);
    expect(parser.isSupported("test.txt")).toBe(true);
    expect(parser.isSupported("test.xlsm")).toBe(true);
    expect(parser.isSupported("test.json")).toBe(true);
    expect(parser.isSupported("test.docx")).toBe(true);
    expect(parser.isSupported("test.pptx")).toBe(true);
    expect(parser.isSupported("test.pdf")).toBe(false);
    expect(parser.isSupported("test.doc")).toBe(false);
    expect(parser.isSupported("test.ppt")).toBe(false);
  });

  it("should parse an XLSX workbook through Open XML", async () => {
    const tmpPath = path.join(os.tmpdir(), `test-${Date.now()}.xlsx`);
    const zip = new JSZip();
    zip.file("xl/workbook.xml", `
      <workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheets><sheet name="People" sheetId="1" r:id="rId1"/></sheets>
      </workbook>
    `);
    zip.file("xl/_rels/workbook.xml.rels", `
      <Relationships>
        <Relationship Id="rId1" Target="worksheets/sheet1.xml"/>
      </Relationships>
    `);
    zip.file("xl/sharedStrings.xml", `
      <sst>
        <si><t>Name</t></si>
        <si><t>Age</t></si>
        <si><t>Ada</t></si>
      </sst>
    `);
    zip.file("xl/worksheets/sheet1.xml", `
      <worksheet>
        <dimension ref="A1:B2"/>
        <sheetData>
          <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>
          <row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2"><v>36</v></c></row>
        </sheetData>
      </worksheet>
    `);
    fs.writeFileSync(tmpPath, await zip.generateAsync({ type: "nodebuffer" }));

    try {
      const chunks = await parser.parseAsync(tmpPath);
      expect(chunks).toHaveLength(1);
      expect(chunks[0].metadata.sheetName).toBe("People");
      expect(chunks[0].metadata.tableRange).toBe("A1:B2");
      expect(chunks[0].metadata.headers).toEqual(["Name", "Age"]);
      expect(chunks[0].metadata.rowCount).toBe(1);
      expect(chunks[0].content).toContain("Ada | 36");
    } finally {
      fs.unlinkSync(tmpPath);
    }
  });

  it("should parse a CSV file correctly", () => {
    const tmpPath = path.join(os.tmpdir(), `test-${Date.now()}.csv`);
    fs.writeFileSync(tmpPath, "姓名,年龄,城市\n张三,28,北京\n李四,35,上海\n", "utf-8");

    try {
      const chunks = parser.parse(tmpPath);
      expect(chunks).toHaveLength(1);
      expect(chunks[0].metadata.headers).toEqual(["姓名", "年龄", "城市"]);
      expect(chunks[0].content).toContain("张三");
      expect(chunks[0].content).toContain("李四");
    } finally {
      fs.unlinkSync(tmpPath);
    }
  });

  it("should parse a JSON file into searchable paths", () => {
    const tmpPath = path.join(os.tmpdir(), `test-${Date.now()}.json`);
    fs.writeFileSync(tmpPath, JSON.stringify({
      formula: {
        name: "区域汇总公式",
        target: "汇总表!B2",
      },
      fields: ["区域", "销售额"],
    }), "utf-8");

    try {
      const chunks = parser.parse(tmpPath);
      expect(chunks).toHaveLength(1);
      expect(chunks[0].sourceType).toBe("json");
      expect(chunks[0].content).toContain("$.formula.name: 区域汇总公式");
      expect(chunks[0].content).toContain("$.fields[1]: 销售额");
    } finally {
      fs.unlinkSync(tmpPath);
    }
  });

  it("should parse DOCX body text through Open XML", async () => {
    const tmpPath = path.join(os.tmpdir(), `test-${Date.now()}.docx`);
    const zip = new JSZip();
    zip.file("word/document.xml", `
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p><w:r><w:t>区域汇总公式说明</w:t></w:r></w:p>
          <w:p><w:r><w:t>公式应该写入汇总表锚点单元格。</w:t></w:r></w:p>
        </w:body>
      </w:document>
    `);
    fs.writeFileSync(tmpPath, await zip.generateAsync({ type: "nodebuffer" }));

    try {
      const chunks = await parser.parseAsync(tmpPath);
      expect(chunks).toHaveLength(1);
      expect(chunks[0].sourceType).toBe("docx");
      expect(chunks[0].content).toContain("区域汇总公式说明");
      expect(chunks[0].content).toContain("汇总表锚点单元格");
    } finally {
      fs.unlinkSync(tmpPath);
    }
  });

  it("should parse PPTX slide text through Open XML", async () => {
    const tmpPath = path.join(os.tmpdir(), `test-${Date.now()}.pptx`);
    const zip = new JSZip();
    zip.file("ppt/slides/slide1.xml", `
      <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:cSld><p:spTree><p:sp><p:txBody>
          <a:p><a:r><a:t>知识库演示页</a:t></a:r></a:p>
          <a:p><a:r><a:t>支持提取 PPT 文本</a:t></a:r></a:p>
        </p:txBody></p:sp></p:spTree></p:cSld>
      </p:sld>
    `);
    zip.file("ppt/slides/slide2.xml", `
      <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:cSld><p:spTree><p:sp><p:txBody>
          <a:p><a:r><a:t>第二页内容</a:t></a:r></a:p>
        </p:txBody></p:sp></p:spTree></p:cSld>
      </p:sld>
    `);
    fs.writeFileSync(tmpPath, await zip.generateAsync({ type: "nodebuffer" }));

    try {
      const chunks = await parser.parseAsync(tmpPath);
      expect(chunks).toHaveLength(2);
      expect(chunks[0].sourceType).toBe("pptx");
      expect(chunks[0].content).toContain("知识库演示页");
      expect(chunks[1].content).toContain("第二页内容");
    } finally {
      fs.unlinkSync(tmpPath);
    }
  });

  it("should parse a markdown file", () => {
    const tmpPath = path.join(os.tmpdir(), `test-${Date.now()}.md`);
    fs.writeFileSync(tmpPath, "# 标题\n\n这是内容。\n\n## 子标题\n\n子内容。\n", "utf-8");

    try {
      const chunks = parser.parse(tmpPath);
      expect(chunks).toHaveLength(1);
      expect(chunks[0].content).toContain("# 标题");
      expect(chunks[0].sourceType).toBe("md");
    } finally {
      fs.unlinkSync(tmpPath);
    }
  });

  it("should parse a plain text file", () => {
    const tmpPath = path.join(os.tmpdir(), `test-${Date.now()}.txt`);
    fs.writeFileSync(tmpPath, "第一行\n第二行\n第三行\n", "utf-8");

    try {
      const chunks = parser.parse(tmpPath);
      expect(chunks).toHaveLength(1);
      expect(chunks[0].content).toContain("第一行");
    } finally {
      fs.unlinkSync(tmpPath);
    }
  });

  it("should throw for unsupported file types", () => {
    expect(() => parser.parse("test.pdf")).toThrow("不支持");
    expect(() => parser.parse("test.exe")).toThrow("不支持");
  });
});

// ============================================================
// TextChunker 测试
// ============================================================

describe("TextChunker", () => {
  let TextChunker: typeof import("./textChunker").TextChunker;
  let chunker: any;

  beforeEach(async () => {
    const mod = await import("./textChunker");
    TextChunker = mod.TextChunker;
    chunker = new TextChunker(100, 3.5); // 小上限便于测试
  });

  it("should keep small content as single chunk", () => {
    const result = chunker.chunk([
      {
        content: "短文本",
        sourcePath: "/test/file.txt",
        sourceName: "file.txt",
        sourceType: "txt",
        metadata: {},
      },
    ]);

    expect(result).toHaveLength(1);
    expect(result[0].content).toBe("短文本");
    expect(result[0].index).toBe(0);
  });

  it("should estimate tokens correctly", () => {
    const tokens = chunker.estimateTokens("hello world test");
    expect(tokens).toBeGreaterThan(0);
  });

  it("should split tabular data with header preserved in each chunk", () => {
    const header = "列A | 列B | 列C";
    // 生成足够的数据行以强制分块
    const dataLines = Array.from({ length: 300 }, (_, i) => `${i}A | ${i}B | ${i}C`);
    const content = [header, ...dataLines].join("\n");

    // 使用极小的块限制来强制分块：最大 30 tokens ≈ 150 字符
    const smallChunker = new TextChunker(30, 5);
    const result = smallChunker.chunk([
      {
        content,
        sourcePath: "/test/data.csv",
        sourceName: "data.csv",
        sourceType: "csv",
        metadata: { headers: ["列A", "列B", "列C"] },
      },
    ]);

    // 应该被分成多个块，每块都含表头
    expect(result.length).toBeGreaterThan(1);
    for (const chunk of result) {
      expect(chunk.content).toContain("列A | 列B | 列C"); // 表头
    }
  });

  it("should split markdown by headings", () => {
    const sections = Array.from({ length: 5 }, (_, i) =>
      `## 第${i + 1}节\n\n这是第${i + 1}节的内容描述。`
    );
    const content = sections.join("\n\n");

    const result = chunker.chunk([
      {
        content,
        sourcePath: "/test/doc.md",
        sourceName: "doc.md",
        sourceType: "md",
        metadata: {},
      },
    ]);

    expect(result.length).toBeGreaterThanOrEqual(1);
  });
});

// ============================================================
// KnowledgeIndexer 测试（集成）
// ============================================================

describe("KnowledgeIndexer", () => {
  let SqliteStore: typeof import("./sqliteStore").SqliteStore;
  let EmbeddingService: typeof import("./embeddingService").EmbeddingService;
  let KnowledgeIndexer: typeof import("./knowledgeIndexer").KnowledgeIndexer;
  let store: any;
  let embedder: any;
  let indexer: any;

  beforeEach(async () => {
    const mod0 = await import("./sqliteStore");
    const mod1 = await import("./embeddingService");
    const mod2 = await import("./knowledgeIndexer");

    SqliteStore = mod0.SqliteStore;
    EmbeddingService = mod1.EmbeddingService;
    KnowledgeIndexer = mod2.KnowledgeIndexer;

    store = new SqliteStore(":memory:");
    await store.init();

    embedder = new EmbeddingService({
      provider: "openai",
      baseUrl: "https://api.openai.com",
      apiKey: "sk-test",
    });

    // Mock embedder to return fixed vectors
    embedder.embedBatch = async (texts: string[]) =>
      texts.map(() => [0.1, 0.2, 0.3]);

    embedder.embed = async (text: string) => [0.1, 0.2, 0.3];

    indexer = new KnowledgeIndexer(store, embedder);
  });

  it("should index a CSV file and create entries", async () => {
    const tmpPath = path.join(os.tmpdir(), `test-index-${Date.now()}.csv`);
    fs.writeFileSync(tmpPath, "商品,价格,库存\n苹果,5.0,100\n香蕉,3.0,200\n", "utf-8");

    try {
      const result = await indexer.indexFile(tmpPath);

      expect(result.success).toBe(true);
      expect(result.entryCount).toBeGreaterThan(0);
      expect(store.countEntries()).toBe(result.entryCount);

      // 验证来源记录已创建
      const source = store.getSource(tmpPath);
      expect(source).not.toBeNull();
      expect(source!.entryCount).toBe(result.entryCount);
      expect(source!.fileHash).toBeTruthy();
    } finally {
      fs.unlinkSync(tmpPath);
    }
  });

  it("should keep keyword-searchable entries when embedding is unavailable", async () => {
    const tmpPath = path.join(os.tmpdir(), `test-keyword-only-${Date.now()}.csv`);
    fs.writeFileSync(tmpPath, "主题,说明\n区域汇总公式,区域汇总公式应该只写到锚点单元格\n", "utf-8");
    embedder.embedBatch = vi.fn(async () => {
      throw new Error("Embedding API 请求失败 (404)");
    });

    try {
      const result = await indexer.indexFile(tmpPath);

      expect(result.success).toBe(true);
      expect(result.entryCount).toBeGreaterThan(0);
      const entries = store.getEntriesBySource(tmpPath);
      expect(entries[0].embedding).toBeNull();
      expect(store.searchByKeyword(["区域汇总公式"], 5)).toHaveLength(1);
    } finally {
      fs.unlinkSync(tmpPath);
    }
  });

  it("should skip unchanged files in incremental mode", async () => {
    const tmpPath = path.join(os.tmpdir(), `test-skip-${Date.now()}.csv`);
    fs.writeFileSync(tmpPath, "a,b\n1,2\n", "utf-8");

    try {
      // 首次索引
      const r1 = await indexer.indexFile(tmpPath);
      expect(r1.success).toBe(true);
      expect(r1.entryCount).toBeGreaterThan(0);

      // 再次索引（未变更）
      const r2 = await indexer.indexFile(tmpPath);
      expect(r2.success).toBe(true);
      // skipUnchanged=true 时，跳过未变更文件，返回现有条数
      expect(r2.entryCount).toBeGreaterThan(0);

      // 条目数应保持不变（没有被重复添加）
      expect(store.countEntries()).toBe(r1.entryCount);
    } finally {
      fs.unlinkSync(tmpPath);
    }
  });

  it("should reindex unchanged files when embedding profile changes", async () => {
    const tmpPath = path.join(os.tmpdir(), `test-profile-${Date.now()}.csv`);
    fs.writeFileSync(tmpPath, "a,b\n1,2\n", "utf-8");

    try {
      const r1 = await indexer.indexFile(tmpPath);
      expect(r1.success).toBe(true);
      expect(r1.entryCount).toBeGreaterThan(0);

      embedder.getProfile = () => ({
        provider: "qwen",
        model: "text-embedding-v2",
        dimensions: 3,
      });
      embedder.embedBatch = vi.fn(async (texts: string[]) =>
        texts.map(() => [0.4, 0.5, 0.6])
      );

      const r2 = await indexer.indexFile(tmpPath);
      expect(r2.success).toBe(true);
      expect(embedder.embedBatch).toHaveBeenCalled();

      const entries = store.getEntriesBySource(tmpPath);
      expect(entries).toHaveLength(r2.entryCount);
      for (const entry of entries) {
        expect(entry.embeddingProvider).toBe("qwen");
        expect(entry.embeddingModel).toBe("text-embedding-v2");
        expect(entry.embeddingDimensions).toBe(3);
      }
    } finally {
      fs.unlinkSync(tmpPath);
    }
  });

  it("should handle non-existent files gracefully", async () => {
    const result = await indexer.indexFile("/nonexistent/file.xlsx");
    expect(result.success).toBe(false);
    expect(result.error).toContain("不存在");
  });

  it("should delete source and all its entries", async () => {
    const tmpPath = path.join(os.tmpdir(), `test-del-${Date.now()}.csv`);
    fs.writeFileSync(tmpPath, "x,y\n1,2\n", "utf-8");

    try {
      await indexer.indexFile(tmpPath);
      const countBefore = store.countEntries();
      expect(countBefore).toBeGreaterThan(0);

      await indexer.deleteSource(tmpPath);
      expect(store.countEntries()).toBe(0);
      expect(store.getSource(tmpPath)).toBeNull();
    } finally {
      fs.unlinkSync(tmpPath);
    }
  });
});

// ============================================================
// Retriever 测试
// ============================================================

describe("Retriever", () => {
  let SqliteStore: typeof import("./sqliteStore").SqliteStore;
  let EmbeddingService: typeof import("./embeddingService").EmbeddingService;
  let Retriever: typeof import("./retriever").Retriever;
  let store: any;
  let embedder: any;
  let retriever: any;

  beforeEach(async () => {
    const mod0 = await import("./sqliteStore");
    const mod1 = await import("./embeddingService");
    const mod2 = await import("./retriever");

    SqliteStore = mod0.SqliteStore;
    EmbeddingService = mod1.EmbeddingService;
    Retriever = mod2.Retriever;

    store = new SqliteStore(":memory:");
    await store.init();

    embedder = new EmbeddingService({
      provider: "openai",
      baseUrl: "https://api.openai.com",
      apiKey: "sk-test",
    });

    // 插入测试数据
    store.bulkInsert([
      {
        id: "r-id-1",
        source: "workbook",
        sourcePath: "/test/sales.xlsx",
        sourceName: "sales.xlsx",
        sourceType: "xlsx",
        chunkIndex: 0,
        content: "销售数据_2024年各月销售额统计表",
        metadata: { sheetName: "月度数据" },
        embedding: [1, 0, 0],
        embeddingProvider: "openai",
        embeddingModel: "text-embedding-3-small",
        embeddingDimensions: 1536,
        indexedAt: Date.now(),
        tokenCount: 10,
      },
      {
        id: "r-id-2",
        source: "document",
        sourcePath: "/test/report.md",
        sourceName: "report.md",
        sourceType: "md",
        chunkIndex: 0,
        content: "财务报表分析报告_2024年度",
        metadata: {},
        embedding: [0.5, 0.5, 0],
        embeddingProvider: "openai",
        embeddingModel: "text-embedding-3-small",
        embeddingDimensions: 1536,
        indexedAt: Date.now(),
        tokenCount: 8,
      },
      {
        id: "r-id-3",
        source: "document",
        sourcePath: "/test/notes.txt",
        sourceName: "notes.txt",
        sourceType: "txt",
        chunkIndex: 0,
        content: "项目进度跟踪_2024年Q3",
        metadata: {},
        embedding: [0, 1, 0],
        embeddingProvider: "openai",
        embeddingModel: "text-embedding-3-small",
        embeddingDimensions: 1536,
        indexedAt: Date.now(),
        tokenCount: 6,
      },
    ]);

    embedder.embed = async (text: string) => {
      if (text.includes("销售")) return [1, 0, 0];
      if (text.includes("财务")) return [0.5, 0.5, 0];
      return [0, 0, 1];
    };

    retriever = new mod2.Retriever(store, embedder, {
      candidateCount: 10,
      defaultTopK: 3,
      minScore: 0.1,
    });
  });

  it("should return relevant results by vector similarity", async () => {
    const results = await retriever.search({ text: "销售", topK: 3 });
    expect(results.length).toBeGreaterThanOrEqual(1);
    expect(results[0].entry.content).toContain("销售");
  });

  it("should fall back to keyword search when query embedding fails", async () => {
    embedder.embed = vi.fn(async () => {
      throw new Error("Embedding API 请求失败 (404)");
    });

    const results = await retriever.search({ text: "区域销售数据应该怎么汇总", topK: 3 });

    expect(results.length).toBeGreaterThanOrEqual(1);
    expect(results[0].entry.content).toContain("销售");
    expect(results[0].score).toBe(0);
    expect(retriever.formatForToolResult(results)).toContain("关键词匹配");
  });

  it("should fall back to keyword search when current embedding profile has no vector matches", async () => {
    embedder.getProfile = () => ({
      provider: "qwen",
      model: "text-embedding-v2",
      dimensions: 1536,
    });

    const results = await retriever.search({
      text: "财务报表",
      topK: 3,
      sourceFilter: ["document"],
    });

    expect(results.map((r: KnowledgeResult) => r.entry.id)).toEqual(["r-id-2"]);
    expect(results[0].score).toBe(0);
  });

  it("should filter by source type", async () => {
    const results = await retriever.search({
      text: "销售",
      topK: 3,
      sourceFilter: ["document"],
    });

    // 应该只返回 document 类型的结果
    for (const r of results) {
      expect(r.entry.source).toBe("document");
    }
  });

  it("should filter by low score threshold", async () => {
    const strictRetriever = new (await import("./retriever")).Retriever(
      store, embedder, { minScore: 0.9 }
    );

    const results = await strictRetriever.search({ text: "销售", topK: 3 });
    // 只有完全匹配的才能超过 0.9
    for (const r of results) {
      expect(r.score).toBeGreaterThanOrEqual(0.9);
    }
  });

  it("should format results for prompt injection", async () => {
    const results = await retriever.search({ text: "销售", topK: 3 });
    const formatted = retriever.formatForPrompt(results);

    expect(formatted).toContain("相关知识");
    expect(formatted).toContain("sales.xlsx");
  });

  it("should format results for tool output", async () => {
    const results = await retriever.search({ text: "销售", topK: 3 });
    const formatted = retriever.formatForToolResult(results);

    expect(formatted).toContain("销售");
    expect(formatted).toContain("相关度");
  });

  it("should return empty for formatForPrompt when no results", () => {
    const formatted = retriever.formatForPrompt([]);
    expect(formatted).toBe("");
  });
});
