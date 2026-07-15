/**
 * 知识检索器
 *
 * 两段式召回：
 * 1. 粗筛：关键词 LIKE 匹配，快速找出候选集
 * 2. 精排：向量余弦相似度排序，取 Top-K
 *
 * 同时支持纯关键词搜索（无嵌入时的降级方案）
 * 和纯向量搜索（关键词匹配不到时）。
 */

import type { KnowledgeEntry, KnowledgeQuery, KnowledgeResult, KnowledgeSourceType } from "./types";
import { SqliteStore } from "./sqliteStore";
import { EmbeddingService, type EmbeddingProfile } from "./embeddingService";

const TOOL_RESULT_EXCERPT_CHARS = 1200;

// ============================================================
// 检索选项
// ============================================================

export interface RetrieverOptions {
  /** 粗筛候选集大小（默认 50） */
  candidateCount?: number;
  /** 精排返回 Top-K（默认 5） */
  defaultTopK?: number;
  /** 最低相似度阈值（默认 0.3） */
  minScore?: number;
}

// ============================================================
// Retriever
// ============================================================

export class Retriever {
  private store: SqliteStore;
  private embedder: EmbeddingService;
  private options: Required<RetrieverOptions>;

  constructor(store: SqliteStore, embedder: EmbeddingService, options?: RetrieverOptions) {
    this.store = store;
    this.embedder = embedder;
    this.options = {
      candidateCount: 50,
      defaultTopK: 5,
      minScore: 0.3,
      ...options,
    };
  }

  /**
   * 执行两段式召回
   *
   * 流程：
   * 1. 将查询文本生成向量嵌入
   * 2. 先用关键词粗筛（可选）
   * 3. 用向量精排
   *
   * @param query - 查询参数
   * @returns 排序后的检索结果
   */
  async search(query: KnowledgeQuery): Promise<KnowledgeResult[]> {
    const topK = query.topK || this.options.defaultTopK;
    const minScore = query.minScore ?? this.options.minScore;

    // 1. 生成查询向量
    let queryVector: number[];
    try {
      queryVector = await this.embedder.embed(query.text);
    } catch {
      return this.searchByKeywordFallback(query, topK);
    }

    // 2. 向量搜索
    const filter: {
      sourceFilter?: string[];
      pathFilter?: string[];
      embeddingProfile?: EmbeddingProfile;
    } = {
      embeddingProfile: this.embedder.getProfile(),
    };
    if (query.sourceFilter && query.sourceFilter.length > 0) {
      filter.sourceFilter = query.sourceFilter;
    }
    if (query.pathFilter && query.pathFilter.length > 0) {
      filter.pathFilter = query.pathFilter;
    }

    // 先用向量搜索获取更多候选，再过滤
    const candidateCount = Math.max(topK * 5, this.options.candidateCount);
    let results = this.store.searchByVector(queryVector, candidateCount, filter);

    // 3. 过滤低分结果
    results = results.filter((r) => r.score >= minScore);

    if (results.length === 0) {
      return this.searchByKeywordFallback(query, topK);
    }

    // 4. 取 Top-K
    return results.slice(0, topK);
  }

  /**
   * 仅关键词搜索（降级方案）
   *
   * 当嵌入服务不可用时使用。
   */
  searchByKeywords(
    keywords: string[],
    topK?: number,
    filter?: { sourceFilter?: KnowledgeSourceType[]; pathFilter?: string[] },
  ): KnowledgeEntry[] {
    const k = topK || this.options.defaultTopK;
    return this.store.searchByKeyword(keywords, k, filter);
  }

  /**
   * 为 system prompt 注入准备格式化文本
   *
   * 将检索结果转换为文本摘要，用于注入到 system prompt 中。
   */
  formatForPrompt(results: KnowledgeResult[]): string {
    if (results.length === 0) return "";

    const lines: string[] = [
      "## 相关知识（用于归纳与校验，不是现成答案）",
      "先从当前任务提炼输入/输出形状、业务键、约束和变换链，再组合下面的知识；不要按相似案例直接替换区域地址。",
    ];
    const seen = new Set<string>();

    for (const result of results) {
      const entry = result.entry;
      const sourceKey = `${entry.sourcePath}:${entry.metadata?.sheetName || ""}`;

      if (!seen.has(sourceKey)) {
        seen.add(sourceKey);
        const sheetInfo = entry.metadata?.sheetName ? ` → ${entry.metadata.sheetName}` : "";
        lines.push(`\n📄 ${entry.sourceName}${sheetInfo}`);
      }

      const excerpt = entry.content.slice(0, 600).replace(/\n/g, " ");
      lines.push(`  - ${excerpt}${entry.content.length > 600 ? "…" : ""}`);
    }

    return lines.join("\n");
  }

  /**
   * 为 knowledge.search 工具准备格式化文本
   *
   * 更详细的格式化，包含来源路径和相似度分数。
   */
  formatForToolResult(results: KnowledgeResult[]): string {
    if (results.length === 0) {
      return "知识库中未找到相关信息。";
    }

    const lines: string[] = [
      `找到 ${results.length} 条相关知识片段。`,
      "使用要求：先归纳当前任务的问题结构和完整变换链，再综合这些片段形成方案；案例内容只用于验证局部技巧或兼容性，禁止直接套用最相似案例。",
    ];

    for (let i = 0; i < results.length; i++) {
      const r = results[i];
      const entry = r.entry;
      const sheetInfo = entry.metadata?.sheetName ? ` (Sheet: ${entry.metadata.sheetName})` : "";
      const scoreLabel = r.score > 0 ? `相关度: ${(r.score * 100).toFixed(1)}%` : "关键词匹配";

      lines.push(`\n### 知识片段 ${i + 1}`);
      lines.push(`来源: [${entry.sourceName}]${sheetInfo} (${scoreLabel})`);
      lines.push(`   路径: ${entry.sourcePath}`);
      lines.push(
        `内容:\n${entry.content.slice(0, TOOL_RESULT_EXCERPT_CHARS)}` +
          `${entry.content.length > TOOL_RESULT_EXCERPT_CHARS ? "…" : ""}`,
      );
    }

    return lines.join("\n");
  }

  /** 更新检索器配置 */
  updateOptions(options: Partial<RetrieverOptions>): void {
    Object.assign(this.options, options);
  }

  private searchByKeywordFallback(query: KnowledgeQuery, topK: number): KnowledgeResult[] {
    const keywords = extractSearchKeywords(query.text);
    const entries = this.searchByKeywords(keywords, topK, {
      sourceFilter: query.sourceFilter,
      pathFilter: query.pathFilter,
    });
    return entries.map((entry) => ({ entry, score: 0 }));
  }
}

function extractSearchKeywords(text: string): string[] {
  const normalized = text.replace(/\s+/g, " ").trim();
  if (!normalized) return [];

  const keywords: string[] = [];
  const add = (value: string) => {
    const item = value.trim();
    if (item.length >= 2 && !keywords.includes(item)) {
      keywords.push(item);
    }
  };

  add(normalized.slice(0, 80));

  for (const part of normalized.split(/[^\p{L}\p{N}_]+/u)) {
    add(part);
    if (/[\p{Script=Han}]/u.test(part) && part.length > 2) {
      for (let size = Math.min(6, part.length); size >= 2; size--) {
        for (let start = 0; start <= part.length - size; start++) {
          add(part.slice(start, start + size));
          if (keywords.length >= 40) return keywords;
        }
      }
    }
    if (keywords.length >= 40) break;
  }

  return keywords;
}
