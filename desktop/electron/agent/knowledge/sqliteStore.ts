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
  KnowledgeSource,
  KnowledgeResult,
} from "./types";
import type { EmbeddingProfile } from "./embeddingService";
import { ensureSourceSummaries } from "./sqliteSourceSummaries";
import { cosineSimilarity, entryToRow, rowToEntry, rowToSource } from "./sqliteStoreRows";
import { initKnowledgeTables } from "./sqliteStoreSchema";
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

    initKnowledgeTables(this.db);
  }

  // ============================================================
  // 写入操作
  // ============================================================

  /** 插入单条知识条目 */
  insertEntry(entry: KnowledgeEntry): void {
    const row = entryToRow(entry);
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
        const row = entryToRow(item);
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

    return row ? rowToEntry(row) : null;
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
        const score = cosineSimilarity(queryVec, entryVec);
        if (score > 0) {
          results.push({ entry: rowToEntry(row), score });
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
  searchByKeyword(
    keywords: string[],
    topK: number,
    filter?: { sourceFilter?: string[]; pathFilter?: string[] }
  ): KnowledgeEntry[] {
    if (keywords.length === 0) return [];

    const seen = new Set<string>();
    const results: KnowledgeEntry[] = [];

    for (const kw of keywords) {
      const where: string[] = ["content LIKE ?"];
      const params: any[] = [`%${kw}%`];
      if (filter?.sourceFilter && filter.sourceFilter.length > 0) {
        where.push(`source IN (${filter.sourceFilter.map(() => "?").join(",")})`);
        params.push(...filter.sourceFilter);
      }
      if (filter?.pathFilter && filter.pathFilter.length > 0) {
        where.push(`source_path IN (${filter.pathFilter.map(() => "?").join(",")})`);
        params.push(...filter.pathFilter);
      }
      params.push(topK);

      const rows = this.db
        .prepare(`SELECT * FROM knowledge_entries WHERE ${where.join(" AND ")} LIMIT ?`)
        .all(...params) as Record<string, any>[];

      for (const row of rows) {
        if (!seen.has(row.id)) {
          seen.add(row.id);
          results.push(rowToEntry(row));
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
    ensureSourceSummaries(this.db);
    const rows = this.db
      .prepare("SELECT * FROM knowledge_sources ORDER BY last_indexed DESC")
      .all() as Record<string, any>[];

    return rows.map((r) => rowToSource(r));
  }

  /** 获取指定来源记录 */
  getSource(sourcePath: string): KnowledgeSource | null {
    const row = this.db
      .prepare("SELECT * FROM knowledge_sources WHERE source_path = ?")
      .get(sourcePath) as Record<string, any> | undefined;

    return row ? rowToSource(row) : null;
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

    return rows.map((r) => rowToEntry(r));
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

}
