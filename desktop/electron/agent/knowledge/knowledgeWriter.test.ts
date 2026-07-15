import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { afterEach, describe, expect, it } from "vitest";

import type { KnowledgeEntry, KnowledgeSource } from "./types";
import { KnowledgeWriter } from "./knowledgeWriter";

describe("KnowledgeWriter", () => {
  const tempDirs: string[] = [];

  afterEach(() => {
    for (const dir of tempDirs.splice(0)) {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });

  it("writes model notes as searchable knowledge entries", async () => {
    const notesDir = fs.mkdtempSync(path.join(os.tmpdir(), "knowledge-writer-"));
    tempDirs.push(notesDir);
    const inserted: KnowledgeEntry[][] = [];
    const sources: KnowledgeSource[] = [];
    const writer = new KnowledgeWriter(
      {
        bulkInsert: (entries) => inserted.push(entries),
        upsertSource: (source) => sources.push(source),
      },
      {
        embedBatch: async (texts) => texts.map(() => [0.1, 0.2, 0.3]),
      },
      { notesDir },
    );

    const result = await writer.writeNote({
      title: "字段口径",
      content: "销售额字段必须扣除退款金额。",
      tags: ["销售", "口径"],
    });

    expect(result.entryCount).toBe(1);
    expect(result.sourceName).toMatch(/^note-\d{14}-[a-f0-9-]+/);
    expect(fs.existsSync(result.sourcePath)).toBe(true);
    expect(fs.readFileSync(result.sourcePath, "utf8")).toContain("销售额字段必须扣除退款金额。");
    expect(inserted[0][0]).toMatchObject({
      source: "note",
      sourceType: "md",
      sourcePath: result.sourcePath,
      metadata: expect.objectContaining({ title: "字段口径", tags: ["销售", "口径"] }),
      embedding: [0.1, 0.2, 0.3],
    });
    expect(sources[0]).toMatchObject({
      sourcePath: result.sourcePath,
      sourceType: "md",
      entryCount: 1,
    });
  });

  it("writes keyword-only notes when embedding is unavailable", async () => {
    const notesDir = fs.mkdtempSync(path.join(os.tmpdir(), "knowledge-writer-"));
    tempDirs.push(notesDir);
    const inserted: KnowledgeEntry[][] = [];
    const writer = new KnowledgeWriter(
      {
        bulkInsert: (entries) => inserted.push(entries),
        upsertSource: () => undefined,
      },
      {
        embedBatch: async () => {
          throw new Error("Embedding API 请求失败 (404)");
        },
      },
      { notesDir },
    );

    const result = await writer.writeNote({
      title: "区域汇总公式",
      content: "区域汇总公式应该只写到锚点单元格。",
    });

    expect(result.entryCount).toBe(1);
    expect(fs.existsSync(result.sourcePath)).toBe(true);
    expect(inserted[0][0]).toMatchObject({
      source: "note",
      embedding: null,
      embeddingProvider: undefined,
      embeddingModel: undefined,
    });
    expect(inserted[0][0].content).toContain("区域汇总公式");
  });
  it("replaces an existing text knowledge source and rebuilds its index", async () => {
    const notesDir = fs.mkdtempSync(path.join(os.tmpdir(), "knowledge-writer-"));
    tempDirs.push(notesDir);
    const sourcePath = path.join(notesDir, "formula-rules.md");
    fs.writeFileSync(sourcePath, "# Old rules\n\nUse fixed sample coordinates.", "utf8");
    const inserted: KnowledgeEntry[][] = [];
    const deletedSources: string[] = [];
    const sources: KnowledgeSource[] = [];
    const writer = new KnowledgeWriter(
      {
        bulkInsert: (entries) => inserted.push(entries),
        upsertSource: (source) => sources.push(source),
        deleteSource: (pathToDelete) => deletedSources.push(pathToDelete),
      },
      {
        embedBatch: async (texts) => texts.map(() => [0.4, 0.5, 0.6]),
      },
      { notesDir },
    );

    const result = await writer.updateSource({
      sourcePath,
      operation: "replace",
      title: "Formula rules",
      content: "Formulas must stay dynamic when the source range grows.",
      tags: ["formula"],
    });

    const fileContent = fs.readFileSync(sourcePath, "utf8");
    expect(result.sourcePath).toBe(sourcePath);
    expect(fileContent).toContain("Formulas must stay dynamic");
    expect(fileContent).not.toContain("fixed sample coordinates");
    expect(deletedSources).toEqual([sourcePath]);
    expect(inserted[0][0]).toMatchObject({
      source: "note",
      sourcePath,
      sourceName: "formula-rules.md",
      sourceType: "md",
      embedding: [0.4, 0.5, 0.6],
      metadata: expect.objectContaining({
        title: "Formula rules",
        tags: ["formula"],
        origin: "model",
      }),
    });
    expect(sources[0]).toMatchObject({
      sourcePath,
      sourceName: "formula-rules.md",
      sourceType: "md",
      entryCount: result.entryCount,
    });
  });

  it("appends to an existing text knowledge source and rebuilds its index", async () => {
    const notesDir = fs.mkdtempSync(path.join(os.tmpdir(), "knowledge-writer-"));
    tempDirs.push(notesDir);
    const sourcePath = path.join(notesDir, "invoice-fields.txt");
    fs.writeFileSync(sourcePath, "Invoice number maps to invoice_no.", "utf8");
    const inserted: KnowledgeEntry[][] = [];
    const writer = new KnowledgeWriter(
      {
        bulkInsert: (entries) => inserted.push(entries),
        upsertSource: () => undefined,
        deleteSource: () => undefined,
      },
      {
        embedBatch: async (texts) => texts.map(() => [0.7]),
      },
      { notesDir },
    );

    await writer.updateSource({
      sourcePath,
      operation: "append",
      content: "Total amount maps to amount_total.",
    });

    const text = fs.readFileSync(sourcePath, "utf8");
    expect(text).toContain("Invoice number maps to invoice_no.");
    expect(text).toContain("Total amount maps to amount_total.");
    expect(inserted[0][0]).toMatchObject({
      source: "note",
      sourcePath,
      sourceType: "txt",
    });
    expect(inserted[0][0].content).toContain("amount_total");
  });

  it("deletes knowledge source index without deleting the source file", async () => {
    const notesDir = fs.mkdtempSync(path.join(os.tmpdir(), "knowledge-writer-"));
    tempDirs.push(notesDir);
    const sourcePath = path.join(notesDir, "keep-file.md");
    fs.writeFileSync(sourcePath, "# Keep file", "utf8");
    const deletedSources: string[] = [];
    const writer = new KnowledgeWriter(
      {
        bulkInsert: () => undefined,
        upsertSource: () => undefined,
        deleteSource: (pathToDelete) => deletedSources.push(pathToDelete),
      },
      {
        embedBatch: async () => [],
      },
      { notesDir },
    );

    const result = await writer.deleteSource({ sourcePath });

    expect(result).toEqual({ sourcePath });
    expect(deletedSources).toEqual([sourcePath]);
    expect(fs.existsSync(sourcePath)).toBe(true);
  });
});
