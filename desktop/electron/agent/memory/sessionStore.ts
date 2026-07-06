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
  type Turn,
  type Thread,
  type ThreadMetadata,
  type TurnStatus,
  type RolloutLine,
  type RolloutItem,
  type TokenUsage,
  mergeTokenUsage,
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
import { readUsageSummaryFromRolloutFiles, type TurnStats } from "./sessionUsageStats";

const readFile = promisify(fs.readFile);
const readdir = promisify(fs.readdir);
const stat = promisify(fs.stat);
const unlink = promisify(fs.unlink);

// ============================================================
// 路径管理
// ============================================================

/** 获取会话存储根目录 */
export function getDefaultSessionsRoot(): string {
  const appData = process.env.APPDATA || path.join(process.env.USERPROFILE || "C:\\", "AppData", "Roaming");
  return path.join(appData, "excel-ai-assistant", "sessions");
}

/** 根据日期生成会话文件路径 */
function getSessionFilePath(sessionsRoot: string, threadId: ThreadId): string {
  const now = new Date();
  const year = now.getFullYear().toString();
  const month = (now.getMonth() + 1).toString().padStart(2, "0");
  const day = now.getDate().toString().padStart(2, "0");

  const dir = path.join(sessionsRoot, year, month, day);
  const timestamp = now.toISOString().replace(/:/g, "-").replace(/\.\d+Z$/, "");
  const filename = `rollout-${timestamp}-${threadId}.jsonl`;

  return path.join(dir, filename);
}

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

  constructor(sessionsRoot?: string, rolloutWriter = new AsyncRolloutWriter()) {
    this.sessionsRoot = sessionsRoot || getDefaultSessionsRoot();
    this.rolloutWriter = rolloutWriter;
  }

  /** 设置数据库 rollout 投影写入器；JSONL 仍保留为兼容审计副本。 */
  setRolloutEventSink(sink: RolloutEventSink | null): void {
    this.rolloutEventSink = sink;
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

  /** 解析 JSONL 内容为 Thread 对象 */
  private parseRolloutContent(content: string, threadId: ThreadId): Thread {
    const lines = content.split("\n").filter((line) => line.trim());
    let metadata: ThreadMetadata = {
      threadId,
      preview: "",
      modelProvider: "unknown",
      createdAt: Date.now(),
      updatedAt: Date.now(),
    };
    const turnsMap = new Map<TurnId, Turn>();
    /** 按时间顺序记录所有遇到的 turnId，用于压缩后清理 */
    const turnOrder: TurnId[] = [];
    let currentTurnId: TurnId | null = null;

    for (const line of lines) {
      try {
        const parsed: RolloutLine = JSON.parse(line);
        const { item } = parsed;

        switch (item.type) {
          case "session_meta": {
            metadata = {
              ...metadata,
              threadId: item.meta.id,
              modelProvider: item.meta.modelProvider,
              model: item.meta.model,
              createdAt: new Date(item.meta.timestamp).getTime(),
              folderId: item.meta.folderId,
            };
            break;
          }
          case "turn_context": {
            currentTurnId = item.turnId;
            if (!turnsMap.has(item.turnId)) {
              turnsMap.set(item.turnId, {
                turnId: item.turnId,
                threadId: metadata.threadId,
                status: "in_progress",
                items: [],
                startedAt: Date.now(),
              });
              turnOrder.push(item.turnId);
            }
            break;
          }
          case "turn_item": {
            const turnId = item.turnId;
            if (!turnsMap.has(turnId)) {
              turnsMap.set(turnId, {
                turnId,
                threadId: metadata.threadId,
                status: "in_progress",
                items: [],
                startedAt: Date.now(),
              });
              turnOrder.push(turnId);
            }
            const turn = turnsMap.get(turnId)!;
            turn.items.push(item.item);

            // 更新元数据
            if (item.item.type === "user_message" && !metadata.preview) {
              metadata.preview = item.item.content.slice(0, 100);
            }
            if (item.item.type === "assistant_message") {
              metadata.updatedAt = item.item.timestamp;
            }
            break;
          }
          case "turn_usage": {
            if (!turnsMap.has(item.turnId)) {
              turnsMap.set(item.turnId, {
                turnId: item.turnId,
                threadId: metadata.threadId,
                status: "in_progress",
                items: [],
                startedAt: Date.now(),
              });
              turnOrder.push(item.turnId);
            }
            const turn = turnsMap.get(item.turnId)!;
            turn.tokenUsage = item.usage;
            metadata.totalTokenUsage = metadata.totalTokenUsage
              ? mergeTokenUsage(metadata.totalTokenUsage, item.usage)
              : item.usage;
            break;
          }
          case "compacted": {
            // 压缩记录：线程级别的历史替代点
            // 1. 保存 replacementHistory 到 metadata，供 AgentLoop 恢复时使用
            metadata.compactedHistory = item.replacementHistory;
            // 2. 清除压缩点之前的所有 turns（它们已被摘要替代）
            //    保留压缩点之后仍在 turnsMap 中的 turns
            for (const tId of turnOrder) {
              turnsMap.delete(tId);
            }
            turnOrder.length = 0;
            break;
          }
        }
      } catch {
        // 跳过解析失败的行（参考 Codex 的容错处理）
        continue;
      }
    }

    // 确定每个 Turn 的最终状态
    const turns = Array.from(turnsMap.values());
    for (const turn of turns) {
      if (turn.status === "in_progress") {
        // 如果最后一轮是 in_progress，说明是中断的
        const hasFinalMessage = turn.items.some(
          (item) => item.type === "assistant_message" && item.phase === "final"
        );
        turn.status = hasFinalMessage ? "completed" : "interrupted";
      }
    }
    const lastTurn = turns.length > 0 ? turns[turns.length - 1] : undefined;
    if (lastTurn) {
      metadata.lastTurnStatus = lastTurn.status;
    }

    return {
      metadata,
      turns,
    };
  }

  // ----------------------------------------------------------
  // 列出所有会话（参考 Codex list_threads）
  // ----------------------------------------------------------

  async listThreads(): Promise<ThreadMetadata[]> {
    await this.flushRolloutWrites();

    const threads: ThreadMetadata[] = [];
    await this.scanDirectory(this.sessionsRoot, threads);
    // 按更新时间降序排列
    threads.sort((a, b) => b.updatedAt - a.updatedAt);
    return threads;
  }

  private async scanDirectory(dir: string, results: ThreadMetadata[]): Promise<void> {
    let entries: fs.Dirent[];
    try {
      entries = await readdir(dir, { withFileTypes: true });
    } catch {
      return;
    }

    for (const entry of entries) {
      const fullPath = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        await this.scanDirectory(fullPath, results);
      } else if (entry.name.endsWith(".jsonl") && entry.name.startsWith("rollout-")) {
        // 快速读取第一行获取元数据
        try {
          const thread = await this.loadThreadByPath(fullPath);
          if (thread) {
            results.push(thread.metadata);
          }
        } catch {
          // 跳过损坏的文件
        }
      }
    }
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
    // 简化实现：在 JSONL 末尾追加一条 session_meta
    // 加载时会合并最后一条 session_meta
    if (patch.modelProvider || patch.model || patch.name || patch.folderId !== undefined) {
      // 读取当前元数据以保留未修改的字段
      const current = await this.loadThread(threadId);
      const currentMeta = current?.metadata;

      await this.appendRolloutItems(threadId, [
        {
          type: "session_meta",
          meta: {
            id: threadId,
            timestamp: new Date().toISOString(),
            modelProvider: patch.modelProvider || currentMeta?.modelProvider || "unknown",
            model: patch.model !== undefined ? patch.model : currentMeta?.model,
            folderId: patch.folderId !== undefined ? patch.folderId : currentMeta?.folderId,
          },
        },
      ]);
    }
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
    const files = await this.findAllRolloutFiles();
    for (const file of files) {
      if (file.includes(threadId)) {
        this.rolloutPathCache.set(threadId, file);
        return file;
      }
    }
    return null;
  }

  /** 查找所有 Rollout 文件 */
  private async findAllRolloutFiles(): Promise<string[]> {
    const files: string[] = [];
    await this.collectRolloutFiles(this.sessionsRoot, files);
    return files;
  }

  private async collectRolloutFiles(dir: string, files: string[]): Promise<void> {
    let entries: fs.Dirent[];
    try {
      entries = await readdir(dir, { withFileTypes: true });
    } catch {
      return;
    }

    for (const entry of entries) {
      const fullPath = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        await this.collectRolloutFiles(fullPath, files);
      } else if (entry.name.endsWith(".jsonl")) {
        files.push(fullPath);
      }
    }
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

    const files = await this.findAllRolloutFiles();
    return readUsageSummaryFromRolloutFiles(files);
  }
}
