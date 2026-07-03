import type {
  ThreadId,
  ThreadRuntimeSnapshot,
  ThreadRuntimeStatus,
} from "../../shared/types";

export const DEFAULT_THREAD_IDLE_UNLOAD_MS = 30 * 60 * 1000;

export interface ThreadStateManagerOptions {
  idleUnloadMs?: number;
}

/**
 * 线程运行态管理。
 *
 * 关联模块：
 * - agentLoop.ts: 根据状态决定是否空闲卸载 activeThread。
 * - threadLifecycle.ts: 负责创建/恢复线程，本模块只记录运行态，不读写 rollout。
 */
export class ThreadStateManager {
  private status: ThreadRuntimeStatus = "not_loaded";
  private threadId: ThreadId | undefined;
  private lastActiveAt: number | undefined;
  private unloadedAt: number | undefined;
  private readonly idleUnloadMs: number;

  constructor(options: ThreadStateManagerOptions = {}) {
    this.idleUnloadMs = options.idleUnloadMs ?? DEFAULT_THREAD_IDLE_UNLOAD_MS;
  }

  markLoaded(threadId: ThreadId, now = Date.now()): void {
    this.status = "active";
    this.threadId = threadId;
    this.lastActiveAt = now;
    this.unloadedAt = undefined;
  }

  markRunning(threadId: ThreadId, now = Date.now()): void {
    this.status = "running";
    this.threadId = threadId;
    this.lastActiveAt = now;
    this.unloadedAt = undefined;
  }

  markIdle(threadId: ThreadId, now = Date.now()): void {
    this.status = "active";
    this.threadId = threadId;
    this.lastActiveAt = now;
  }

  markUnloaded(now = Date.now()): void {
    if (!this.threadId) return;
    this.status = "unloaded";
    this.lastActiveAt = now;
    this.unloadedAt = now;
  }

  clear(): void {
    this.status = "not_loaded";
    this.threadId = undefined;
    this.lastActiveAt = undefined;
    this.unloadedAt = undefined;
  }

  shouldUnload(now = Date.now()): boolean {
    if (this.status !== "active" || !this.threadId || this.lastActiveAt === undefined) {
      return false;
    }
    if (this.idleUnloadMs <= 0) return false;
    return now - this.lastActiveAt >= this.idleUnloadMs;
  }

  getSnapshot(): ThreadRuntimeSnapshot {
    return {
      status: this.status,
      threadId: this.threadId,
      lastActiveAt: this.lastActiveAt,
      unloadedAt: this.unloadedAt,
      idleUnloadMs: this.idleUnloadMs,
    };
  }
}
