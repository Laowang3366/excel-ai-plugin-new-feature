/**
 * 运行中用户输入队列。
 *
 * 关联模块：
 * - agentLoop.ts: 运行中收到的新输入先入队，当前 turn 结束后继续处理。
 * - interaction/ipcAgentHandlers.ts: IPC 层只负责入队，不直接操作队列内部状态。
 */

import type { AgentTurnCallbacks, AgentTurnInput } from "../../shared/types";

export interface QueuedTurnInput {
  input: AgentTurnInput;
  callbacks: AgentTurnCallbacks;
}

export class InputQueue {
  private readonly items: QueuedTurnInput[] = [];
  private readonly maxItems: number;

  constructor(maxItems = 32) {
    this.maxItems = Math.max(1, Math.floor(maxItems));
  }

  enqueue(item: QueuedTurnInput): number {
    if (this.items.length >= this.maxItems) {
      throw new Error(
        `运行中输入队列已满（${this.maxItems} 条），请等待当前会话处理一部分内容后再发送`,
      );
    }
    this.items.push(item);
    return this.items.length;
  }

  dequeue(): QueuedTurnInput | undefined {
    return this.items.shift();
  }

  clear(): number {
    const count = this.items.length;
    this.items.length = 0;
    return count;
  }

  size(): number {
    return this.items.length;
  }
}
