/**
 * AI 流式响应收集器
 *
 * 从 AgentLoop.runAgentLoop 中提取的流式响应处理逻辑。
 * 负责消费 AI streamChat 的事件流，收集文本、推理、工具调用等数据。
 */

import {
  type AIStreamEvent,
  type ChatMessage,
  type ReasoningMode,
  type TokenUsage,
} from "../../providers/aiClient";
import { type AgentTurnCallbacks, type ToolCallItem, type TurnItem } from "../../shared/types";
import { isRetriableAIRequestError } from "./aiRequestRetry";

// ============================================================
// 类型
// ============================================================

/** 流式参数 */
export interface StreamParams {
  messages: ChatMessage[];
  tools: import("../../shared/types").ToolDefinition[];
  systemPrompt: string;
  maxTokens: number;
  reasoningMode: ReasoningMode;
  signal?: AbortSignal;
  /**
   * 采样轮次 ID（一次 turn 内的 round 序号）。
   * 同一轮内所有 streamDelta 携带同一 roundId，
   * 让前端能在跨轮切换时主动重置 streaming buffers。
   */
  roundId?: number;
}

/** 流式收集结果 */
export interface StreamResult {
  assistantContent: string;
  reasoningContent: string[];
  reasoningSummary: string[];
  toolCalls: ToolCallInfo[];
  finishReason: string;
  usage: TokenUsage | undefined;
  /** 流式阶段已创建的 tool_call item，key 为 toolCallId */
  pendingToolCallItems: Map<string, ToolCallItem>;
}

/** 工具调用基本信息 */
export interface ToolCallInfo {
  id: string;
  name: string;
  arguments: string;
}

// ============================================================
// 收集器
// ============================================================

/**
 * 收集 AI 流式响应的全部事件
 *
 * @param streamIterable - AI streamChat 返回的异步迭代器
 * @param callbacks - 事件回调（用于 onStreamDelta、onEvent 等）
 * @returns 收集完成的 StreamResult
 *
 * @throws AbortError - 如果用户中断（由调用方处理）
 * @throws Error - AI 流式错误（由调用方处理）
 */
export async function collectStreamEvents(
  streamIterable: AsyncIterable<AIStreamEvent>,
  callbacks: AgentTurnCallbacks,
  roundId?: number,
): Promise<StreamResult> {
  const result: StreamResult = {
    assistantContent: "",
    reasoningContent: [],
    reasoningSummary: [],
    toolCalls: [],
    finishReason: "",
    usage: undefined,
    pendingToolCallItems: new Map(),
  };

  // 跟踪流式参数拼接
  const pendingToolCallArgs = new Map<string, string>();
  const pendingToolCallNames = new Map<string, string>();

  for await (const event of streamIterable) {
    switch (event.type) {
      case "text_delta": {
        result.assistantContent += event.delta;
        callbacks.onStreamDelta?.(event.delta, "assistant_message", roundId);
        break;
      }
      case "reasoning_delta": {
        result.reasoningContent.push(event.delta);
        callbacks.onStreamDelta?.(event.delta, "reasoning", roundId);
        break;
      }
      case "reasoning_summary_delta": {
        result.reasoningSummary.push(event.delta);
        break;
      }
      case "tool_call_begin": {
        pendingToolCallArgs.set(event.toolCallId, "");
        pendingToolCallNames.set(event.toolCallId, event.toolName);

        // 创建 tool_call item
        const toolCallItem: ToolCallItem = {
          type: "tool_call",
          id: event.toolCallId,
          toolName: event.toolName,
          arguments: {},
          status: "pending",
          timestamp: Date.now(),
        };
        result.pendingToolCallItems.set(event.toolCallId, toolCallItem);

        // 立即发出 tool_call item_started，让前端第一时间看到工具开始
        // 避免 emitStreamResultItems 延迟期间 UI 仍显示思考态
        // (emitStreamResultItems 会再来一次 item_started，前端是幂等的)
        callbacks.onEvent({ type: "item_started", item: toolCallItem });
        break;
      }
      case "tool_call_delta": {
        const existing = pendingToolCallArgs.get(event.toolCallId) || "";
        pendingToolCallArgs.set(event.toolCallId, existing + event.delta);
        break;
      }
      case "tool_call_end": {
        // 更新已创建的 tool_call item 的参数
        const existingItem = result.pendingToolCallItems.get(event.toolCallId);
        if (existingItem) {
          try {
            existingItem.arguments = JSON.parse(event.arguments || "{}");
          } catch {
            existingItem.arguments = { _raw: event.arguments };
          }
          // 通知前端更新参数
          callbacks.onEvent({ type: "item_updated", item: existingItem });
        }

        result.toolCalls.push({
          id: event.toolCallId,
          name: event.toolName,
          arguments: event.arguments,
        });
        break;
      }
      case "usage": {
        result.usage = event.usage;
        break;
      }
      case "done": {
        result.finishReason = event.finishReason;
        break;
      }
      case "error": {
        if (isRetriableAIRequestError(new Error(event.error))) {
          throw new Error(event.error);
        }
        // 错误事件直接封装为 TurnItem 返回，由调用方决定如何处理
        const errorItem: TurnItem = {
          type: "error",
          id: `error-${Date.now()}`,
          message: event.error,
          timestamp: Date.now(),
        };
        // 通过特殊的 error 字段传递给调用方
        result.assistantContent = ""; // 清空，标记有错误
        (result as any).errorItem = errorItem;
        return result;
      }
    }
  }

  return result;
}
