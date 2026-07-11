/**
 * 会话存储 — 参考 Codex 的 thread-store 和 rollout 架构
 *
 * 核心功能：
 * 1. 会话持久化：将对话记录以 JSONL 格式保存到本地文件
 * 2. 会话恢复：从文件中加载历史记录，支持中断后继续
 * 3. 会话列表：列出所有历史会话
 * 4. Rollout 记录：参考 Codex 的 rollout recorder，异步写入事件
 *
 * 文件结构：
 *   ~/AppData/excel-ai-assistant/sessions/
 *     └── YYYY/MM/DD/
 *         └── rollout-YYYY-MM-DDThh-mm-ss-{threadId}.jsonl
 */

import * as fs from "fs";
import * as path from "path";
import { promisify } from "util";
import {
  type ThreadId,
  type TurnId,
  type TurnItem,
  type Thread,
  type ThreadMetadata,
  type TurnStatus,
  type RolloutLine,
  type RolloutItem,
  type TokenUsage,
  generateThreadId,
  generateTurnId,
} from "../shared/types";
import { clampNumber } from "../shared/numberLimits";
import { AsyncRolloutWriter } from "./rolloutWriter";
import {
  searchCompressedRolloutMatches,
  spawnRolloutCompressionWorker,
  type CompressedRolloutSearchMatch,
  type RolloutCompressionWorkerOptions,
  type RolloutCompressionWorkerResult,
} from "./rolloutArchive";
import type { RuntimeRolloutSearchMatch } from "./stateRuntimeTypes";
import { parseRolloutContent as parseSessionRolloutContent } from "./sessionRolloutParser";
import { readUsageSummaryFromRolloutFiles, type TurnStats } from "./sessionUsageStats";
import {
  findAllRolloutFiles,
  getDefaultSessionsRoot,
  getSessionFilePath,
  scanThreadMetadata,
} from "./sessionStoreFiles";

const readFile = promisify(fs.readFile);
const unlink = promisify(fs.unlink);

export { getDefaultSessionsRoot };

export interface RolloutEventSink {
  appendRolloutItems(threadId: ThreadId, items: RolloutItem[]): Promise<void>;
  searchRolloutMatches?(query: string, options?: { limit?: number }): Promise<RuntimeRolloutSearchMatch[]>;
}

export type SessionRolloutSearchMatch = RuntimeRolloutSearchMatch | CompressedRolloutSearchMatch;

// ============================================================
// SessionStore — 会话存储与恢复
// ============================================================

export class SessionStore {
  private sessionsRoot: string;
  private rolloutWriter: AsyncRolloutWriter;
  private rolloutEventSink: RolloutEventSink | null = null;
  private writesSuspendedReason: string | null = null;

  constructor(sessionsRoot?: string, rolloutWriter = new AsyncRolloutWriter()) {
    this.sessionsRoot = sessionsRoot || getDefaultSessionsRoot();
    this.rolloutWriter = rolloutWriter;
  }

  /** 设置数据库 rollout 投影写入器；JSONL 仍保留为兼容审计副本。 */
  setRolloutEventSink(sink: RolloutEventSink | null): void {
    this.rolloutEventSink = sink;
  }

  suspendWrites(reason = "会话存储正在维护"): void {
    this.writesSuspendedReason = reason;
  }

  resumeWrites(): void {
    this.writesSuspendedReason = null;
  }

  // ----------------------------------------------------------
  // 创建新会话
  // ----------------------------------------------------------

  async createThread(modelProvider: string, model?: string, folderId?: string): Promise<Thread> {
    const threadId = generateThreadId();
    const now = Date.now();

    const metadata: ThreadMetadata = {
      threadId,
      preview: "",
      modelProvider,
      model,
      createdAt: now,
      updatedAt: now,
      folderId,
    };

    const thread: Thread = {
      metadata,
      turns: [],
    };

    // 写入 session_meta 作为 JSONL 的第一行
    await this.appendRolloutItems(threadId, [
      {
        type: "session_meta",
        meta: {
          id: threadId,
          timestamp: new Date(now).toISOString(),
          modelProvider,
          model,
          folderId,
        },
      },
    ]);

    return thread;
  }

  // ----------------------------------------------------------
  // 追加 Rollout 条目
  // ----------------------------------------------------------

  async appendRolloutItems(threadId: ThreadId, items: RolloutItem[]): Promise<void> {
    if (items.length === 0) return;
    if (this.writesSuspendedReason) {
      throw new Error(this.writesSuspendedReason);
    }

    if (this.rolloutEventSink) {
      await this.rolloutEventSink.appendRolloutItems(threadId, items);
    }

    const filePath = this.getRolloutPath(threadId);

    // 构造 JSONL 行
    const lines = items.map((item) => {
      const line: RolloutLine = {
        timestamp: new Date().toISOString(),
        item,
      };
      return JSON.stringify(line) + "\n";
    });

    await this.rolloutWriter.enqueue(filePath, lines);
  }

  /** 等待所有已排队的 rollout 写入落盘。读取、退出或测试前调用。 */
  async flushRolloutWrites(): Promise<void> {
    await this.rolloutWriter.flush();
  }

  /** 启动冷 rollout JSONL zstd 压缩 worker。 */
  async spawnRolloutCompressionWorker(
    options: Omit<RolloutCompressionWorkerOptions, "sessionsRoot">
  ): Promise<RolloutCompressionWorkerResult> {
    await this.flushRolloutWrites();
    return spawnRolloutCompressionWorker({
      ...options,
      sessionsRoot: this.sessionsRoot,
    });
  }

  /** 搜索数据库投影和压缩 JSONL 归档中的 rollout 内容。 */
  async searchRolloutMatches(
    query: string,
    options: { limit?: number } = {}
  ): Promise<SessionRolloutSearchMatch[]> {
    await this.flushRolloutWrites();
    const limit = clampNumber(options.limit, { fallback: 20, min: 1, max: 100 });
    const dbMatches = this.rolloutEventSink?.searchRolloutMatches
      ? await this.rolloutEventSink.searchRolloutMatches(query, { limit })
      : [];
    if (dbMatches.length >= limit) return dbMatches.slice(0, limit);

    const archiveMatches = await searchCompressedRolloutMatches({
      sessionsRoot: this.sessionsRoot,
      query,
      limit: limit - dbMatches.length,
    });
    return [...dbMatches, ...archiveMatches].slice(0, limit);
  }

  // ----------------------------------------------------------
  // 追加 TurnItem 到指定 Turn
  // ----------------------------------------------------------

  async appendTurnItem(threadId: ThreadId, turnId: TurnId, item: TurnItem): Promise<void> {
    await this.appendRolloutItems(threadId, [
      { type: "turn_item", turnId, item },
    ]);
  }

  async appendTurnUsage(threadId: ThreadId, turnId: TurnId, usage: TokenUsage): Promise<void> {
    await this.appendRolloutItems(threadId, [
      { type: "turn_usage", turnId, usage },
    ]);
  }

  // ----------------------------------------------------------
  // 从文件恢复会话（参考 Codex load_rollout_items）
  // ----------------------------------------------------------

  async loadThread(threadId: ThreadId): Promise<Thread | null> {
    await this.flushRolloutWrites();

    // 先尝试缓存路径
    let filePath: string | null | undefined = this.rolloutPathCache.get(threadId);

    // 缓存未命中时，扫描磁盘查找实际文件
    if (!filePath) {
      filePath = await this.findRolloutPath(threadId);
    }

    if (!filePath) return null;

    try {
      const content = await readFile(filePath, "utf-8");
      return this.parseRolloutContent(content, threadId);
    } catch (err: any) {
      if (err.code === "ENOENT") return null;
      throw err;
    }
  }

  /** 通过文件路径加载会话 */
  async loadThreadByPath(filePath: string): Promise<Thread | null> {
    await this.flushRolloutWrites();

    try {
      const content = await readFile(filePath, "utf-8");
      // 从文件路径中提取 threadId
      const filename = path.basename(filePath, ".jsonl");
      const threadIdMatch = filename.match(/thread-(.+)$/);
      const threadId = threadIdMatch ? `thread-${threadIdMatch[1]}` : filename;
      return this.parseRolloutContent(content, threadId);
    } catch (err: any) {
      if (err.code === "ENOENT") return null;
      throw err;
    }
  }

  /** Parse JSONL content into a Thread object. */
  private parseRolloutContent(content: string, threadId: ThreadId): Thread {
    return parseSessionRolloutContent(content, threadId);
  }

  // ----------------------------------------------------------
  // 列出所有会话（参考 Codex list_threads）
  // ----------------------------------------------------------

  async listThreads(): Promise<ThreadMetadata[]> {
    await this.flushRolloutWrites();

    const threads = await scanThreadMetadata(this.sessionsRoot, (filePath) =>
      this.loadThreadByPath(filePath)
    );
    // 按更新时间降序排列
    threads.sort((a, b) => b.updatedAt - a.updatedAt);
    return threads;
  }

  // ----------------------------------------------------------
  // 删除会话
  // ----------------------------------------------------------

  async deleteThread(threadId: ThreadId): Promise<boolean> {
    await this.flushRolloutWrites();

    // 优先使用缓存路径，否则扫描磁盘查找实际文件
    let filePath: string | null | undefined = this.rolloutPathCache.get(threadId);
    if (!filePath) {
      filePath = await this.findRolloutPath(threadId);
    }
    if (!filePath) return false;

    try {
      await unlink(filePath);
      // 清除缓存
      this.rolloutPathCache.delete(threadId);
      return true;
    } catch {
      return false;
    }
  }

  // ----------------------------------------------------------
  // 更新会话元数据
  // ----------------------------------------------------------

  async updateThreadMetadata(threadId: ThreadId, patch: Partial<ThreadMetadata>): Promise<void> {
    const current = await this.loadThread(threadId);
    if (!current) throw new Error("会话不存在");

    const merged = { ...current.metadata, ...patch, updatedAt: Date.now() };
    await this.appendRolloutItems(threadId, [{
      type: "session_meta",
      meta: {
        id: threadId,
        timestamp: new Date(merged.updatedAt).toISOString(),
        modelProvider: merged.modelProvider,
        model: merged.model,
        name: merged.name ?? null,
        folderId: merged.folderId,
      },
    }]);
  }

  // ----------------------------------------------------------
  // Rollout 路径管理
  // ----------------------------------------------------------

  /** threadId → 文件路径的映射缓存 */
  private rolloutPathCache = new Map<ThreadId, string>();

  /** 注册/获取 Rollout 文件路径 */
  getRolloutPath(threadId: ThreadId): string {
    if (this.rolloutPathCache.has(threadId)) {
      return this.rolloutPathCache.get(threadId)!;
    }
    // 新建时生成路径
    const filePath = getSessionFilePath(this.sessionsRoot, threadId);
    this.rolloutPathCache.set(threadId, filePath);
    return filePath;
  }

  /** 注册已有的 Rollout 路径（恢复会话时使用） */
  registerRolloutPath(threadId: ThreadId, filePath: string): void {
    this.rolloutPathCache.set(threadId, filePath);
  }

  /** 查找已存在的 Rollout 文件 */
  async findRolloutPath(threadId: ThreadId): Promise<string | null> {
    if (this.rolloutPathCache.has(threadId)) {
      return this.rolloutPathCache.get(threadId)!;
    }

    // 扫描目录查找匹配的文件
    const files = await findAllRolloutFiles(this.sessionsRoot);
    for (const file of files) {
      if (file.includes(threadId)) {
        this.rolloutPathCache.set(threadId, file);
        return file;
      }
    }
    return null;
  }

  // ----------------------------------------------------------
  // 获取最近的会话（用于自动恢复）
  // ----------------------------------------------------------

  async findLatestThread(): Promise<ThreadId | null> {
    const threads = await this.listThreads();
    if (threads.length === 0) return null;
    // listThreads 已按 updatedAt 降序排列
    return threads[0].threadId;
  }

  // ----------------------------------------------------------
  // 使用统计聚合（单次扫描，替代 N+1 的 thread:load 调用）
  // ----------------------------------------------------------

  async getUsageSummary(): Promise<TurnStats[]> {
    await this.flushRolloutWrites();

    const files = await findAllRolloutFiles(this.sessionsRoot);
    return readUsageSummaryFromRolloutFiles(files);
  }
}
