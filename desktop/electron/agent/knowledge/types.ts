/**
 * RAG 知识增强层 — 核心类型定义
 *
 * 定义知识条目、来源、查询、结果等核心类型，
 * 供整个 RAG 子系统使用。
 */

// ============================================================
// 知识来源类型
// ============================================================

/** 知识来源分类 */
export type KnowledgeSourceType =
  | "workbook"      // Excel 工作簿（.xlsx/.xlsm/.xlsb）
  | "document"      // 文档（.csv/.md/.txt）
  | "note"          // 用户手写笔记
  | "agents_md";    // AGENTS.md 项目级知识

/** 文件类型 */
export type KnowledgeFileType =
  | "xlsx"
  | "xlsm"
  | "xlsb"
  | "csv"
  | "md"
  | "txt";

// ============================================================
// 知识条目
// ============================================================

/**
 * 知识条目 — 最小的知识单元
 *
 * 一个文件被解析成多个 chunk，每个 chunk 对应一个 KnowledgeEntry。
 * embedding 为 null 表示尚未生成向量嵌入（待处理）。
 */
export interface KnowledgeEntry {
  /** 唯一标识（UUID） */
  id: string;
  /** 来源分类 */
  source: KnowledgeSourceType;
  /** 来源文件的绝对路径 */
  sourcePath: string;
  /** 来源文件名（含扩展名） */
  sourceName: string;
  /** 文件类型 */
  sourceType: KnowledgeFileType;
  /** 分块序号 */
  chunkIndex: number;
  /** 文本内容 */
  content: string;
  /** 元数据（JSON 对象，不同来源有不同字段） */
  metadata: Record<string, unknown>;
  /** 向量嵌入（null 表示未嵌入） */
  embedding: number[] | null;
  embeddingProvider?: string;
  embeddingModel?: string;
  embeddingDimensions?: number;
  /** 索引时间戳（ms） */
  indexedAt: number;
  /** 估算 token 数 */
  tokenCount: number;
}

// ============================================================
// 知识来源摘要
// ============================================================

/** 已索引的知识来源摘要（用于前端展示） */
export interface KnowledgeSource {
  /** 来源文件的绝对路径 */
  sourcePath: string;
  /** 文件名 */
  sourceName: string;
  /** 文件类型 */
  sourceType: KnowledgeFileType;
  /** 条目数 */
  entryCount: number;
  /** 首次索引时间戳 */
  firstIndexed: number;
  /** 最后索引时间戳 */
  lastIndexed: number;
  /** 文件内容哈希（用于增量检测） */
  fileHash: string;
}

// ============================================================
// 查询与结果
// ============================================================

/** 知识库查询参数 */
export interface KnowledgeQuery {
  /** 搜索文本 */
  text: string;
  /** 返回 Top-K 条结果 */
  topK: number;
  /** 按来源类型过滤（可选） */
  sourceFilter?: KnowledgeSourceType[];
  /** 按文件路径过滤（可选） */
  pathFilter?: string[];
  /** 最低相似度分数（0-1，可选） */
  minScore?: number;
}

/** 检索结果 */
export interface KnowledgeResult {
  /** 匹配的知识条目 */
  entry: KnowledgeEntry;
  /** 相似度分数（0-1，越高越相关） */
  score: number;
}

// ============================================================
// 索引结果
// ============================================================

/** 文件索引结果 */
export interface IndexResult {
  /** 来源文件路径 */
  sourcePath: string;
  /** 是否成功 */
  success: boolean;
  /** 错误信息（失败时） */
  error?: string;
  /** 生成的条目数 */
  entryCount: number;
  /** 耗时（ms） */
  durationMs: number;
}

// ============================================================
// 数据库行记录（内部使用）
// ============================================================

/** 数据库中的 knowledge_entries 行（含 embedding JSON 字符串） */
export interface KnowledgeEntryRow {
  id: string;
  source: KnowledgeSourceType;
  source_path: string;
  source_name: string;
  source_type: KnowledgeFileType;
  chunk_index: number;
  content: string;
  metadata: string;       // JSON 字符串
  embedding: string | null; // JSON 浮点数数组字符串
  embedding_provider: string | null;
  embedding_model: string | null;
  embedding_dimensions: number | null;
  indexed_at: number;
  token_count: number;
}
