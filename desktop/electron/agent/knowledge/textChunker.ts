/**
 * 文本分块器
 *
 * 将解析后的原始块按 token 上限进一步拆分，
 * 确保每个 chunk 不超过 512 tokens（约 2000 字符），便于嵌入和检索。
 *
 * 分块策略：
 * - Excel / CSV（表格数据）：按行数拆分，每 100 行一组
 * - Markdown：按 ## / ### 标题层级拆分
 * - 纯文本：按空行分段落，合并小段落
 */

import type { RawChunk } from "./documentParser";

// ============================================================
// 分块结果
// ============================================================

/** 最终的知识块（可直接嵌入） */
export interface TextChunk {
  /** 文本内容 */
  content: string;
  /** 块索引 */
  index: number;
  /** 来源文件路径 */
  sourcePath: string;
  /** 文件名 */
  sourceName: string;
  /** 源类型 */
  sourceType: string;
  /** 元数据 */
  metadata: Record<string, unknown>;
  /** 估算 token 数 */
  tokenCount: number;
}

// ============================================================
// TextChunker
// ============================================================

export class TextChunker {
  /** Token 估算上限 */
  private maxTokens: number;
  /** 每字符≈多少 token（中文约 0.4，英文约 0.25，取保守值） */
  private charsPerToken: number;

  constructor(maxTokens = 512, charsPerToken = 3.5) {
    this.maxTokens = maxTokens;
    this.charsPerToken = charsPerToken;
  }

  /**
   * 将原始块拆分为最终的知识块
   *
   * @param rawChunks - 文档解析器的输出
   * @returns 分块后的 TextChunk 数组
   */
  chunk(rawChunks: RawChunk[]): TextChunk[] {
    const result: TextChunk[] = [];

    for (const raw of rawChunks) {
      const chunks = this.chunkSingle(raw);
      result.push(...chunks);
    }

    return result;
  }

  /**
   * 对单个原始块进行分块
   */
  private chunkSingle(raw: RawChunk): TextChunk[] {
    const tokenEstimate = this.estimateTokens(raw.content);

    // 如果整体不超过上限，直接作为一个块
    if (tokenEstimate <= this.maxTokens) {
      return [
        {
          content: raw.content,
          index: 0,
          sourcePath: raw.sourcePath,
          sourceName: raw.sourceName,
          sourceType: raw.sourceType,
          metadata: raw.metadata,
          tokenCount: tokenEstimate,
        },
      ];
    }

    // 超过上限，按类型分块
    switch (raw.sourceType) {
      case "xlsx":
      case "xlsm":
      case "xlsb":
      case "csv":
        return this.chunkTabular(raw);
      case "md":
        return this.chunkMarkdown(raw);
      default:
        return this.chunkPlainText(raw);
    }
  }

  // ============================================================
  // 表格数据分块（按行拆分）
  // ============================================================

  /**
   * 表格数据分块策略：
   * 1. 固定保留表头行在所有块中
   * 2. 数据行按 ROWS_PER_CHUNK 分组
   * 3. 每块开头重复表头
   */
  private chunkTabular(raw: RawChunk): TextChunk[] {
    const lines = raw.content.split("\n");
    if (lines.length <= 1) {
      return [this.makeChunk(raw, raw.content, 0)];
    }

    // 第一行是表头，保留
    const headerLine = lines[0];
    const dataLines = lines.slice(1);

    // 过滤掉 "...（还有 N 行未展示）" 这类行，它们应该只在最后一块出现
    const summaryLine = dataLines.find((l) => l.startsWith("...（还有"));
    const pureDataLines = dataLines.filter((l) => !l.startsWith("...（还有"));

    const rowsPerChunk = 100;
    const chunks: TextChunk[] = [];

    for (let i = 0; i < pureDataLines.length; i += rowsPerChunk) {
      const chunkLines = pureDataLines.slice(i, i + rowsPerChunk);
      const isLast = i + rowsPerChunk >= pureDataLines.length;

      let content = headerLine + "\n" + chunkLines.join("\n");
      if (isLast && summaryLine) {
        content += "\n" + summaryLine;
      }

      chunks.push(this.makeChunk(raw, content, chunks.length));
    }

    return chunks;
  }

  // ============================================================
  // Markdown 分块（按标题拆分）
  // ============================================================

  /**
   * Markdown 分块策略：
   * 1. 按 ## 或 ### 标题拆分
   * 2. 小段落合并到前一块
   * 3. 大段落单独成块
   */
  private chunkMarkdown(raw: RawChunk): TextChunk[] {
    const lines = raw.content.split("\n");

    // 找到标题行位置
    const headingIndices: number[] = [];
    for (let i = 0; i < lines.length; i++) {
      if (/^#{2,4}\s/.test(lines[i])) {
        headingIndices.push(i);
      }
    }

    // 如果没有标题，按纯文本方式分块
    if (headingIndices.length === 0) {
      return this.chunkPlainText(raw);
    }

    // 按标题分块
    const sections: string[] = [];
    for (let i = 0; i < headingIndices.length; i++) {
      const start = headingIndices[i];
      const end = i + 1 < headingIndices.length ? headingIndices[i + 1] : lines.length;
      sections.push(lines.slice(start, end).join("\n"));
    }

    // 处理每个 section
    const chunks: TextChunk[] = [];
    let currentBuffer = "";
    let currentCount = 0;

    for (const section of sections) {
      const sectionTokens = this.estimateTokens(section);

      if (currentCount + sectionTokens <= this.maxTokens) {
        currentBuffer += (currentBuffer ? "\n\n" : "") + section;
        currentCount += sectionTokens;
      } else {
        if (currentBuffer) {
          chunks.push(this.makeChunk(raw, currentBuffer, chunks.length));
        }
        // 如果单个 section 超过上限，单独成块
        if (sectionTokens > this.maxTokens) {
          chunks.push(this.makeChunk(raw, section, chunks.length));
          currentBuffer = "";
          currentCount = 0;
        } else {
          currentBuffer = section;
          currentCount = sectionTokens;
        }
      }
    }

    if (currentBuffer) {
      chunks.push(this.makeChunk(raw, currentBuffer, chunks.length));
    }

    return chunks;
  }

  // ============================================================
  // 纯文本分块（按段落拆分）
  // ============================================================

  /**
   * 纯文本分块策略：
   * 1. 按空行分段落
   * 2. 小段落合并到不超过上限
   * 3. 超大段落强制截断
   */
  private chunkPlainText(raw: RawChunk): TextChunk[] {
    const paragraphs = raw.content.split(/\n\s*\n/).filter((p) => p.trim());

    const chunks: TextChunk[] = [];
    let currentBuffer = "";
    let currentCount = 0;

    for (const para of paragraphs) {
      const paraTokens = this.estimateTokens(para);

      if (currentCount + paraTokens <= this.maxTokens) {
        currentBuffer += (currentBuffer ? "\n\n" : "") + para;
        currentCount += paraTokens;
      } else {
        if (currentBuffer) {
          chunks.push(this.makeChunk(raw, currentBuffer, chunks.length));
        }
        // 处理超大段落
        if (paraTokens > this.maxTokens) {
          // 硬截断
          const maxChars = Math.floor(this.maxTokens * this.charsPerToken);
          const truncated = para.slice(0, maxChars);
          chunks.push(this.makeChunk(raw, truncated, chunks.length));
          currentBuffer = "";
          currentCount = 0;
        } else {
          currentBuffer = para;
          currentCount = paraTokens;
        }
      }
    }

    if (currentBuffer) {
      chunks.push(this.makeChunk(raw, currentBuffer, chunks.length));
    }

    return chunks.length > 0 ? chunks : [this.makeChunk(raw, raw.content, 0)];
  }

  // ============================================================
  // 工具
  // ============================================================

  /** 创建分块对象 */
  private makeChunk(raw: RawChunk, content: string, index: number): TextChunk {
    return {
      content,
      index,
      sourcePath: raw.sourcePath,
      sourceName: raw.sourceName,
      sourceType: raw.sourceType,
      metadata: raw.metadata,
      tokenCount: this.estimateTokens(content),
    };
  }

  /**
   * 估算文本的 token 数
   *
   * 混合文本（中英文）：按字符数 / charsPerToken 估算。
   * 3.5 字符/token 是对中文（~2）和英文（~4）的折中。
   */
  estimateTokens(text: string): number {
    if (!text) return 0;
    return Math.ceil(text.length / this.charsPerToken);
  }
}
