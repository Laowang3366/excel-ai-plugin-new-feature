export type ConnectionRequestId = string;

/**
 * 待处理中断请求队列。
 *
 * 关联模块：
 * - agentLoop.ts: interrupt() 入队请求 ID，当前 Turn 清理完成后统一 drain。
 * - interaction/ipcAgentHandlers.ts: IPC 中断请求可传入来源 ID，便于后续诊断。
 */
export class PendingInterruptQueue {
  private readonly requestIds: ConnectionRequestId[] = [];

  push(requestId: ConnectionRequestId): void {
    this.requestIds.push(requestId);
  }

  pendingIds(): ConnectionRequestId[] {
    return [...this.requestIds];
  }

  drain(): ConnectionRequestId[] {
    return this.requestIds.splice(0);
  }
}
