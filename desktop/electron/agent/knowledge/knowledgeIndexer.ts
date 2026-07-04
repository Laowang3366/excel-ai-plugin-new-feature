/**
 * 知识索引器
 *
 * 索引编排核心：解析 → 分块 → 嵌入 → 存储。
 * 支持单文件、文件夹递归、增量索引。
 *
 * 使用流程：
 *   const indexer = new KnowledgeIndexer(store, embedder);
 *   const result = await indexer.indexFile("/path/to/file.xlsx");
 */

import * as fs from "fs";
import * as path from "path";
import { createHash, randomUUID } from "crypto";
import type {
  KnowledgeEntry,
  KnowledgeSource,
  KnowledgeSourceType,
  IndexResult,
} from "./types";
import { SqliteStore } from "./sqliteStore";
import { EmbeddingService, type EmbeddingProfile } from "./embeddingService";
import { DocumentParser, type RawChunk } from "./documentParser";
import { TextChunker, type TextChunk } from "./textChunker";

// ============================================================
// 索引选项
// ============================================================

export interface IndexOptions {
  /** 分块最大 token 数（默认 512） */
  maxTokens?: number;
  /** 是否跳过已索引且未变更的文件（增量模式默认 true） */
  skipUnchanged?: boolean;
  /** 嵌入批处理大小（默认 20） */
  batchSize?: number;
  /** 回调函数：进度通知 */
  onProgress?: (current: number, total: number, file: string) => void;
}

// ============================================================
// KnowledgeIndexer
// ============================================================

export class KnowledgeIndexer {
  private store: SqliteStore;
  private embedder: EmbeddingService;
  private parser: DocumentParser;
  private chunker: TextChunker;

  constructor(
    store: SqliteStore,
    embedder: EmbeddingService,
    parser?: DocumentParser,
    chunker?: TextChunker
  ) {
    this.store = store;
    this.embedder = embedder;
    this.parser = parser || new DocumentParser();
    this.chunker = chunker || new TextChunker();
  }

  /**
   * 索引单个文件
   *
   * @param filePath - 文件的绝对路径
   * @param options - 索引选项
   * @returns 索引结果
   */
  async indexFile(
    filePath: string,
    options?: IndexOptions
  ): Promise<IndexResult> {
    const startTime = Date.now();
    const opts: IndexOptions = {
      maxTokens: 512,
      skipUnchanged: true,
      batchSize: 20,
      ...options,
    };

    try {
      // 验证文件存在
      if (!fs.existsSync(filePath)) {
        return {
          sourcePath: filePath,
          success: false,
          error: `文件不存在: ${filePath}`,
          entryCount: 0,
          durationMs: Date.now() - startTime,
        };
      }

      // 增量模式：检查文件哈希
      if (opts.skipUnchanged) {
        const fileHash = this.computeFileHash(filePath);
        const existing = this.store.getSource(filePath);
        if (
          existing
          && existing.fileHash === fileHash
          && this.store.hasSourceEmbeddingProfile(filePath, this.embedder.getProfile())
        ) {
          return {
            sourcePath: filePath,
            success: true,
            entryCount: existing.entryCount,
            durationMs: Date.now() - startTime,
          };
        }
      }

      // 1. 解析
      const rawChunks = await this.parser.parseAsync(filePath);

      // 2. 分块
      const textChunks = this.chunker.chunk(rawChunks);

      if (textChunks.length === 0) {
        return {
          sourcePath: filePath,
          success: true,
          entryCount: 0,
          durationMs: Date.now() - startTime,
        };
      }

      // 3. 生成嵌入向量（批处理）。Embedding 不可用时仍保存文本索引，检索时走关键词兜底。
      const texts = textChunks.map((c) => c.content);
      const embeddingResult = await this.embedChunks(texts);

      // 4. 构建 KnowledgeEntry 并批量写入
      const now = Date.now();
      const sourceType = this.getSourceType(filePath);
      const sourceName = path.basename(filePath);
      const sourceFileType = this.getSourceFileType(filePath);
      const fileHash = this.computeFileHash(filePath);
      const embeddingProfile = embeddingResult.profile;

      // 先删除旧索引
      this.store.deleteSource(filePath);

      // 批量插入
      const entries: KnowledgeEntry[] = [];
      for (let i = 0; i < textChunks.length; i++) {
        const chunk = textChunks[i];
        entries.push({
          id: randomUUID(),
          source: sourceType,
          sourcePath: filePath,
          sourceName,
          sourceType: chunk.sourceType as any,
          chunkIndex: chunk.index,
          content: chunk.content,
          metadata: chunk.metadata,
          embedding: embeddingResult.embeddings[i],
          embeddingProvider: embeddingResult.embeddings[i] ? embeddingProfile?.provider : undefined,
          embeddingModel: embeddingResult.embeddings[i] ? embeddingProfile?.model : undefined,
          embeddingDimensions: embeddingResult.embeddings[i]?.length ?? embeddingProfile?.dimensions,
          indexedAt: now,
          tokenCount: chunk.tokenCount,
        });
      }
      this.store.bulkInsert(entries);

      // 更新来源记录
      this.store.upsertSource({
        sourcePath: filePath,
        sourceName,
        sourceType: sourceFileType as any,
        entryCount: entries.length,
        firstIndexed: now,
        lastIndexed: now,
        fileHash,
      });

      return {
        sourcePath: filePath,
        success: true,
        entryCount: entries.length,
        durationMs: Date.now() - startTime,
      };
    } catch (err: any) {
      return {
        sourcePath: filePath,
        success: false,
        error: err.message,
        entryCount: 0,
        durationMs: Date.now() - startTime,
      };
    }
  }

  /**
   * 递归索引文件夹中的所有支持文件
   *
   * @param folderPath - 文件夹路径
   * @param options - 索引选项
   * @returns 索引结果数组
   */
  async indexFolder(
    folderPath: string,
    options?: IndexOptions
  ): Promise<IndexResult[]> {
    const files = this.collectFiles(folderPath);
    const results: IndexResult[] = [];

    const opts: IndexOptions = {
      maxTokens: 512,
      skipUnchanged: true,
      batchSize: 20,
      ...options,
    };

    const supportedExts = this.parser.getSupportedExtensions();

    // 过滤支持的文件
    const supportedFiles = files.filter((f) => {
      const ext = path.extname(f).toLowerCase().replace(".", "");
      return supportedExts.includes(ext);
    });

    const total = supportedFiles.length;
    if (total === 0) return results;

    for (let i = 0; i < total; i++) {
      const filePath = supportedFiles[i];
      opts.onProgress?.(i + 1, total, path.basename(filePath));
      const result = await this.indexFile(filePath, { ...opts, skipUnchanged: true });
      results.push(result);
    }

    return results;
  }

  /**
   * 删除来源的所有索引
   */
  async deleteSource(sourcePath: string): Promise<void> {
    this.store.deleteSource(sourcePath);
  }

  /**
   * 重建全部索引
   *
   * 扫描所有已注册的来源，重新解析和嵌入。
   */
  async reindexAll(options?: IndexOptions): Promise<IndexResult[]> {
    const sources = this.store.listSources();
    const results: IndexResult[] = [];
    const total = sources.length;

    const opts: IndexOptions = {
      maxTokens: 512,
      skipUnchanged: false, // 重建时强制重新索引
      batchSize: 20,
      ...options,
    };

    for (let i = 0; i < total; i++) {
      const source = sources[i];
      opts.onProgress?.(i + 1, total, source.sourceName);
      const result = await this.indexFile(source.sourcePath, opts);
      results.push(result);
    }

    return results;
  }

  /**
   * 增量索引（只处理新增/变更文件）
   *
   * 对比 file_hash，只索引有变更的文件。
   */
  async incrementalIndex(
    filePaths: string[],
    options?: IndexOptions
  ): Promise<IndexResult[]> {
    const results: IndexResult[] = [];

    for (const fp of filePaths) {
      const result = await this.indexFile(fp, {
        skipUnchanged: true,
        ...options,
      });
      results.push(result);
    }

    return results;
  }

  // ============================================================
  // 内部工具
  // ============================================================

  /** 收集文件夹中所有文件（递归） */
  private collectFiles(dir: string): string[] {
    const files: string[] = [];
    try {
      const entries = fs.readdirSync(dir, { withFileTypes: true });
      for (const entry of entries) {
        const fullPath = path.join(dir, entry.name);
        if (entry.isDirectory()) {
          // 跳过隐藏目录和 node_modules
          if (!entry.name.startsWith(".") && entry.name !== "node_modules") {
            files.push(...this.collectFiles(fullPath));
          }
        } else if (entry.isFile()) {
          files.push(fullPath);
        }
      }
    } catch {
      // 跳过不可读目录
    }
    return files;
  }

  /** 根据路径判断知识来源类型 */
  private getSourceType(filePath: string): KnowledgeSourceType {
    const ext = path.extname(filePath).toLowerCase();
    if ([".xlsx", ".xlsm", ".xlsb"].includes(ext)) return "workbook";
    if (ext === ".md") return "agents_md";
    return "document";
  }

  private getSourceFileType(filePath: string): string {
    return path.extname(filePath).toLowerCase().replace(/^\./, "") || "txt";
  }

  private async embedChunks(texts: string[]): Promise<{
    embeddings: Array<number[] | null>;
    profile?: EmbeddingProfile;
  }> {
    try {
      const embeddings = await this.embedder.embedBatch(texts);
      return {
        embeddings,
        profile: this.embedder.getProfile(),
      };
    } catch {
      return {
        embeddings: texts.map(() => null),
      };
    }
  }

  /** 计算文件哈希（SHA256 前 16 字符） */
  private computeFileHash(filePath: string): string {
    try {
      const content = fs.readFileSync(filePath);
      return createHash("sha256").update(content).digest("hex").slice(0, 16);
    } catch {
      return "";
    }
  }
}
