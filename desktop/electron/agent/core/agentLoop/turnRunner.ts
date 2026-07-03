import {
  generateTurnId,
  type AgentTurnInput,
  type ThreadId,
  type Turn,
  type TurnItem,
} from "../../shared/types";

/**
 * Turn 创建与收尾的小型装配函数。
 *
 * 关联模块：
 * - agentLoop.ts: 单轮运行仍保留在核心循环内，本模块只承载 Turn 数据结构构造。
 * - shared/types: 统一 Turn/TurnItem 形状。
 */
export function createTurn(threadId: ThreadId): Turn {
  return {
    turnId: generateTurnId(),
    threadId,
    status: "in_progress",
    items: [],
    startedAt: Date.now(),
  };
}

export function createUserMessageItem(input: AgentTurnInput): TurnItem {
  return {
    type: "user_message",
    id: `msg-${Date.now()}`,
    content: input.content,
    attachments: input.attachments,
    clientId: input.clientId,
    timestamp: Date.now(),
  };
}

export function completeTurn(turn: Turn): void {
  turn.status = "completed";
  turn.completedAt = Date.now();
}
