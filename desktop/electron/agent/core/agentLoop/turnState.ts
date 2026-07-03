import type { Thread, Turn, TurnItem } from "../../shared/types";

/**
 * AgentLoop 可变运行状态容器。
 *
 * 关联模块：
 * - agentLoop.ts: 通过私有访问器读写当前线程、当前 Turn 和中断状态。
 * - threadLifecycle.ts: 创建/恢复线程后写入 activeThread 与 compactedHistory。
 */
export class TurnState {
  activeThread: Thread | null = null;
  activeTurn: Turn | null = null;
  isRunning = false;
  abortController: AbortController | null = null;
  turnCompletionPromise: Promise<void> | null = null;
  resolveTurnCompletion: (() => void) | null = null;
  pendingFolderId: string | undefined;
  compactedHistory: TurnItem[] | null = null;

  resetForNextThread(folderId?: string): void {
    this.activeThread = null;
    this.activeTurn = null;
    this.compactedHistory = null;
    this.pendingFolderId = folderId;
  }

  consumePendingFolderId(): string | undefined {
    const folderId = this.pendingFolderId;
    this.pendingFolderId = undefined;
    return folderId;
  }
}
