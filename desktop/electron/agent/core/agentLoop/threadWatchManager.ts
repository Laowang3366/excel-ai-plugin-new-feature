import type { ThreadId, ThreadRuntimeSnapshot } from "../../shared/types";

export type ConnectionId = string;
export type ThreadStatusListener = (status: ThreadRuntimeSnapshot) => void;

export interface ThreadWatchSubscription {
  threadId: ThreadId;
  connectionId: ConnectionId;
  close(): void;
}

/**
 * 线程状态订阅管理。
 *
 * 关联模块：
 * - agentLoop.ts: 在线程状态变化时 publish，避免交互层轮询核心状态。
 * - interaction/ipcAgentHandlers.ts: 后续 IPC watch 可用 ConnectionId 跟踪观察者来源。
 */
export class ThreadWatchManager {
  private readonly statuses = new Map<ThreadId, ThreadRuntimeSnapshot>();
  private readonly watchers = new Map<ThreadId, Map<ConnectionId, ThreadStatusListener>>();
  private readonly activeConnections = new Map<ThreadId, Set<ConnectionId>>();

  watch(
    threadId: ThreadId,
    connectionId: ConnectionId,
    listener: ThreadStatusListener,
  ): ThreadWatchSubscription {
    let threadWatchers = this.watchers.get(threadId);
    if (!threadWatchers) {
      threadWatchers = new Map();
      this.watchers.set(threadId, threadWatchers);
    }
    threadWatchers.set(connectionId, listener);

    const currentStatus = this.statuses.get(threadId);
    if (currentStatus) listener(currentStatus);

    return {
      threadId,
      connectionId,
      close: () => {
        this.unwatch(threadId, connectionId);
      },
    };
  }

  publish(status: ThreadRuntimeSnapshot): void {
    if (!status.threadId) return;
    this.statuses.set(status.threadId, status);

    const threadWatchers = this.watchers.get(status.threadId);
    if (!threadWatchers) return;
    for (const listener of threadWatchers.values()) {
      listener(status);
    }
  }

  createActiveGuard(threadId: ThreadId, connectionId: ConnectionId): ThreadWatchActiveGuard {
    let active = this.activeConnections.get(threadId);
    if (!active) {
      active = new Set();
      this.activeConnections.set(threadId, active);
    }
    active.add(connectionId);
    return new ThreadWatchActiveGuard(this, threadId, connectionId);
  }

  getConnectionIds(threadId: ThreadId): ConnectionId[] {
    return [...(this.watchers.get(threadId)?.keys() ?? [])];
  }

  getActiveConnectionIds(threadId: ThreadId): ConnectionId[] {
    return [...(this.activeConnections.get(threadId)?.values() ?? [])];
  }

  private unwatch(threadId: ThreadId, connectionId: ConnectionId): void {
    const threadWatchers = this.watchers.get(threadId);
    if (!threadWatchers) return;
    threadWatchers.delete(connectionId);
    if (threadWatchers.size === 0) {
      this.watchers.delete(threadId);
    }
  }

  releaseActive(threadId: ThreadId, connectionId: ConnectionId): void {
    const active = this.activeConnections.get(threadId);
    if (!active) return;
    active.delete(connectionId);
    if (active.size === 0) {
      this.activeConnections.delete(threadId);
    }
  }
}

export class ThreadWatchActiveGuard {
  private released = false;

  constructor(
    private readonly manager: ThreadWatchManager,
    private readonly threadId: ThreadId,
    private readonly connectionId: ConnectionId,
  ) {}

  dispose(): void {
    if (this.released) return;
    this.released = true;
    this.manager.releaseActive(this.threadId, this.connectionId);
  }
}
