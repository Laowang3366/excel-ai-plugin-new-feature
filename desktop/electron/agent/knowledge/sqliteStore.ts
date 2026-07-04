/**
 * SQLite 存储层 — 基于 Node/Electron 内置 node:sqlite
 *
 * 使用 node:sqlite 管理知识条目的持久化存储。
 * 支持 WAL 模式、FTS5 等完整 SQLite 特性。
 *
 * 设计要点：
 * - 原生 SQLite3，无需 WASM 加载
 * - 自动文件持久化（无需手动 export/save）
 * - WAL 模式支持更好的并发读性能
 * - 向量嵌入以 JSON 数组文本形式存储，余弦相似度在 JS 层计算
 */

import * as path from "path";
import * as fs from "fs";

import type {
  KnowledgeEntry,
  KnowledgeEntryRow,
  KnowledgeSource,
  KnowledgeResult,
} from "./types";
import type { EmbeddingProfile } from "./embeddingService";
import {
  openSqliteDatabase,
  runPragma,
  runSqliteTransaction,
} from "../storage/nodeSqlite";
import type { SqliteDatabase } from "../storage/nodeSqlite";

// ============================================================
// SqliteStore
// ============================================================

export class SqliteStore {
  private db!: SqliteDatabase;
  private dbPath: string;

  constructor(dbPath: string) {
    this.dbPath = dbPath;
  }

  /**
   * 初始化（建表）
   *
   * 保持 async 签名以兼容旧调用方。
   */
  async init(): Promise<void> {
    // :memory: 模式 — 纯内存数据库，不创建文件
    if (this.dbPath !== ":memory:") {
      const dir = path.dirname(this.dbPath);
      if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
      }
    }

    this.db = openSqliteDatabase(this.dbPath);

    // 启用 WAL 模式以提升并发读性能
    runPragma(this.db, "journal_mode = WAL");

    this.initTables();
  }

  /** 创建表（幂等） */
  private initTables(): void {
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS knowledge_entries (
        id          TEXT PRIMARY KEY,
        source      TEXT NOT NULL,
        source_path TEXT NOT NULL,
        source_name TEXT NOT NULL,
        source_type TEXT NOT NULL,
        chunk_index INTEGER DEFAULT 0,
        content     TEXT NOT NULL,
        metadata    TEXT DEFAULT '{}',
        embedding   TEXT,
        embedding_provider TEXT,
        embedding_model TEXT,
        embedding_dimensions INTEGER,
        indexed_at  INTEGER NOT NULL,
        token_count INTEGER DEFAULT 0
      );

      CREATE INDEX IF NOT EXISTS idx_entries_source
        ON knowledge_entries(source);
      CREATE INDEX IF NOT EXISTS idx_entries_source_path
        ON knowledge_entries(source_path);
      CREATE INDEX IF NOT EXISTS idx_entries_indexed_at
        ON knowledge_entries(indexed_at);

      CREATE TABLE IF NOT EXISTS knowledge_sources (
        source_path   TEXT PRIMARY KEY,
        source_name   TEXT NOT NULL,
        source_type   TEXT NOT NULL,
        entry_count   INTEGER DEFAULT 0,
        first_indexed INTEGER NOT NULL,
        last_indexed  INTEGER NOT NULL,
        file_hash     TEXT DEFAULT ''
      );
    `);
    this.migrateEmbeddingProfileColumns();
  }

  private migrateEmbeddingProfileColumns(): void {
    const rows = this.db
      .prepare("PRAGMA table_info(knowledge_entries)")
      .all() as Array<{ name: string }>;
    const columns = new Set(rows.map((row) => row.name));

    if (!columns.has("embedding_provider")) {
      this.db.exec("ALTER TABLE knowledge_entries ADD COLUMN embedding_provider TEXT");
    }
    if (!columns.has("embedding_model")) {
      this.db.exec("ALTER TABLE knowledge_entries ADD COLUMN embedding_model TEXT");
    }
    if (!columns.has("embedding_dimensions")) {
      this.db.exec("ALTER TABLE knowledge_entries ADD COLUMN embedding_dimensions INTEGER");
    }

    this.db.exec(`
      CREATE INDEX IF NOT EXISTS idx_entries_embedding_profile
        ON knowledge_entries(embedding_provider, embedding_model, embedding_dimensions);
    `);
  }

  // ============================================================
  // 写入操作
  // ============================================================

  /** 插入单条知识条目 */
  insertEntry(entry: KnowledgeEntry): void {
    const row = this.entryToRow(entry);
    this.db
      .prepare(
        `INSERT OR REPLACE INTO knowledge_entries
          (id, source, source_path, source_name, source_type,
           chunk_index, content, metadata, embedding,
           embedding_provider, embedding_model, embedding_dimensions,
           indexed_at, token_count)
        VALUES
          (?, ?, ?, ?, ?,
           ?, ?, ?, ?,
           ?, ?, ?,
           ?, ?)`
      )
      .run(
        row.id, row.source, row.source_path, row.source_name, row.source_type,
        row.chunk_index, row.content, row.metadata, row.embedding,
        row.embedding_provider, row.embedding_model, row.embedding_dimensions,
        row.indexed_at, row.token_count
      );
  }

  /** 批量插入知识条目（事务） */
  bulkInsert(entries: KnowledgeEntry[]): void {
    if (entries.length === 0) return;

    const insert = this.db.prepare(
      `INSERT OR REPLACE INTO knowledge_entries
        (id, source, source_path, source_name, source_type,
         chunk_index, content, metadata, embedding,
         embedding_provider, embedding_model, embedding_dimensions,
         indexed_at, token_count)
      VALUES
        (?, ?, ?, ?, ?,
         ?, ?, ?, ?,
         ?, ?, ?,
         ?, ?)`
    );

    const batchInsert = (items: KnowledgeEntry[]) => runSqliteTransaction(this.db, () => {
      for (const item of items) {
        const row = this.entryToRow(item);
        insert.run(
          row.id, row.source, row.source_path, row.source_name, row.source_type,
          row.chunk_index, row.content, row.metadata, row.embedding,
          row.embedding_provider, row.embedding_model, row.embedding_dimensions,
          row.indexed_at, row.token_count
        );
      }
    });

    batchInsert(entries);
  }

  /** 删除单条知识条目 */
  deleteEntry(id: string): void {
    this.db.prepare("DELETE FROM knowledge_entries WHERE id = ?").run(id);
  }

  /** 删除来源下的所有条目和来源记录 */
  deleteSource(sourcePath: string): void {
    const delEntries = this.db.prepare("DELETE FROM knowledge_entries WHERE source_path = ?");
    const delSource = this.db.prepare("DELETE FROM knowledge_sources WHERE source_path = ?");

    const cleanup = (path: string) => runSqliteTransaction(this.db, () => {
      delEntries.run(path);
      delSource.run(path);
    });

    cleanup(sourcePath);
  }

  // ============================================================
  // 查询操作
  // ============================================================

  /** 按 ID 获取单条条目 */
  getEntry(id: string): KnowledgeEntry | null {
    const row = this.db
      .prepare("SELECT * FROM knowledge_entries WHERE id = ?")
      .get(id) as Record<string, any> | undefined;

    return row ? this.rowToEntry(row) : null;
  }

  /**
   * 向量搜索
   *
   * 加载所有含 embedding 的条目，计算余弦相似度后返回 Top-K。
   */
  searchByVector(
    queryVector: number[],
    topK: number,
    filter?: { sourceFilter?: string[]; pathFilter?: string[]; embeddingProfile?: EmbeddingProfile }
  ): KnowledgeResult[] {
    let sql = "SELECT * FROM knowledge_entries WHERE embedding IS NOT NULL";
    const params: any[] = [];
    const expectedProfile = filter?.embeddingProfile;

    if (expectedProfile) {
      sql += " AND embedding_provider = ? AND embedding_model = ? AND embedding_dimensions = ?";
      params.push(expectedProfile.provider, expectedProfile.model, expectedProfile.dimensions);
    }

    if (filter?.sourceFilter && filter.sourceFilter.length > 0) {
      sql += ` AND source IN (${filter.sourceFilter.map(() => "?").join(",")})`;
      params.push(...filter.sourceFilter);
    }
    if (filter?.pathFilter && filter.pathFilter.length > 0) {
      sql += ` AND source_path IN (${filter.pathFilter.map(() => "?").join(",")})`;
      params.push(...filter.pathFilter);
    }

    const rows = this.db.prepare(sql).all(...params) as Record<string, any>[];
    const queryVec = new Float64Array(queryVector);
    const results: KnowledgeResult[] = [];

    for (const row of rows) {
      if (!row.embedding) continue;
      try {
        const entryVec = new Float64Array(JSON.parse(row.embedding));
        const score = this.cosineSimilarity(queryVec, entryVec);
        if (score > 0) {
          results.push({ entry: this.rowToEntry(row), score });
        }
      } catch {
        continue;
      }
    }

    results.sort((a, b) => b.score - a.score);
    return results.slice(0, topK);
  }

  /**
   * 关键词搜索（LIKE 匹配）
   */
  searchByKeyword(keywords: string[], topK: number): KnowledgeEntry[] {
    if (keywords.length === 0) return [];

    const seen = new Set<string>();
    const results: KnowledgeEntry[] = [];

    for (const kw of keywords) {
      const rows = this.db
        .prepare("SELECT * FROM knowledge_entries WHERE content LIKE ? LIMIT ?")
        .all(`%${kw}%`, topK) as Record<string, any>[];

      for (const row of rows) {
        if (!seen.has(row.id)) {
          seen.add(row.id);
          results.push(this.rowToEntry(row));
          if (results.length >= topK) break;
        }
      }
      if (results.length >= topK) break;
    }

    return results.slice(0, topK);
  }

  // ============================================================
  // 来源管理
  // ============================================================

  /** 插入或更新来源记录 */
  upsertSource(source: KnowledgeSource): void {
    this.db
      .prepare(
        `INSERT OR REPLACE INTO knowledge_sources
          (source_path, source_name, source_type, entry_count,
           first_indexed, last_indexed, file_hash)
        VALUES
          (?, ?, ?, ?,
           ?, ?, ?)`
      )
      .run(
        source.sourcePath, source.sourceName, source.sourceType, source.entryCount,
        source.firstIndexed, source.lastIndexed, source.fileHash
      );
  }

  /** 列出所有已索引的来源 */
  listSources(): KnowledgeSource[] {
    const rows = this.db
      .prepare("SELECT * FROM knowledge_sources ORDER BY last_indexed DESC")
      .all() as Record<string, any>[];

    return rows.map((r) => this.rowToSource(r));
  }

  /** 获取指定来源记录 */
  getSource(sourcePath: string): KnowledgeSource | null {
    const row = this.db
      .prepare("SELECT * FROM knowledge_sources WHERE source_path = ?")
      .get(sourcePath) as Record<string, any> | undefined;

    return row ? this.rowToSource(row) : null;
  }

  hasSourceEmbeddingProfile(sourcePath: string, profile: EmbeddingProfile): boolean {
    const row = this.db
      .prepare(
        `SELECT COUNT(*) as count
         FROM knowledge_entries
         WHERE source_path = ?
           AND embedding_provider = ?
           AND embedding_model = ?
           AND embedding_dimensions = ?`
      )
      .get(sourcePath, profile.provider, profile.model, profile.dimensions) as { count: number } | undefined;

    return (row?.count ?? 0) > 0;
  }

  /** 获取指定来源的所有条目 */
  getEntriesBySource(sourcePath: string): KnowledgeEntry[] {
    const rows = this.db
      .prepare("SELECT * FROM knowledge_entries WHERE source_path = ?")
      .all(sourcePath) as Record<string, any>[];

    return rows.map((r) => this.rowToEntry(r));
  }

  // ============================================================
  // 统计
  // ============================================================

  /** 统计总条目数 */
  countEntries(): number {
    const row = this.db
      .prepare("SELECT COUNT(*) as count FROM knowledge_entries")
      .get() as { count: number };

    return row?.count ?? 0;
  }

  // ============================================================
  // 维护
  // ============================================================

  /** 清空所有数据 */
  clearAll(): void {
    this.db.exec("DELETE FROM knowledge_entries");
    this.db.exec("DELETE FROM knowledge_sources");
  }

  /** 执行 VACUUM 回收空间 */
  vacuum(): void {
    this.db.exec("VACUUM");
  }

  /** 关闭数据库连接 */
  close(): void {
    this.db.close();
  }

  /** 获取数据库路径 */
  getDbPath(): string {
    return this.dbPath;
  }

  /** 判断是否已初始化 */
  isInitialized(): boolean {
    return !!this.db;
  }

  // ============================================================
  // 内部工具
  // ============================================================

  /** KnowledgeEntry → 扁平对象（用于绑定参数） */
  private entryToRow(entry: KnowledgeEntry): KnowledgeEntryRow {
    return {
      id: entry.id,
      source: entry.source,
      source_path: entry.sourcePath,
      source_name: entry.sourceName,
      source_type: entry.sourceType,
      chunk_index: entry.chunkIndex,
      content: entry.content,
      metadata: JSON.stringify(entry.metadata),
      embedding: entry.embedding ? JSON.stringify(entry.embedding) : null,
      embedding_provider: entry.embeddingProvider ?? null,
      embedding_model: entry.embeddingModel ?? null,
      embedding_dimensions: entry.embeddingDimensions ?? (entry.embedding ? entry.embedding.length : null),
      indexed_at: entry.indexedAt,
      token_count: entry.tokenCount,
    };
  }

  /** 扁平对象 → KnowledgeEntry */
  private rowToEntry(row: Record<string, any>): KnowledgeEntry {
    return {
      id: row.id,
      source: row.source,
      sourcePath: row.source_path,
      sourceName: row.source_name,
      sourceType: row.source_type,
      chunkIndex: row.chunk_index,
      content: row.content,
      metadata: JSON.parse(row.metadata || "{}"),
      embedding: row.embedding ? JSON.parse(row.embedding) : null,
      embeddingProvider: row.embedding_provider ?? undefined,
      embeddingModel: row.embedding_model ?? undefined,
      embeddingDimensions: row.embedding_dimensions ?? undefined,
      indexedAt: row.indexed_at,
      tokenCount: row.token_count,
    };
  }

  /** 扁平对象 → KnowledgeSource */
  private rowToSource(row: Record<string, any>): KnowledgeSource {
    return {
      sourcePath: row.source_path,
      sourceName: row.source_name,
      sourceType: row.source_type,
      entryCount: row.entry_count,
      firstIndexed: row.first_indexed,
      lastIndexed: row.last_indexed,
      fileHash: row.file_hash,
    };
  }

  /** 余弦相似度 */
  private cosineSimilarity(a: Float64Array, b: Float64Array): number {
    if (a.length !== b.length) return 0;
    const len = a.length;
    let dotProduct = 0;
    let normA = 0;
    let normB = 0;
    for (let i = 0; i < len; i++) {
      dotProduct += a[i] * b[i];
      normA += a[i] * a[i];
      normB += b[i] * b[i];
    }
    const denom = Math.sqrt(normA) * Math.sqrt(normB);
    if (denom === 0) return 0;
    return dotProduct / denom;
  }
}
