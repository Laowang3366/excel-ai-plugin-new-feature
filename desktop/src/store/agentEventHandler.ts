/**
 * agentEventHandler — Agent 事件处理逻辑
 *
 * 从 chatStore.ts 提取，纯函数：接收 AgentEvent → 返回状态 patches。
 * 不依赖 Zustand，便于单元测试。
 */

import type { AgentEvent } from "../electronApi";
import type { ChatState } from "./chatStore";

const FROZEN_REASONING_PREFIX = "streaming-reasoning-";
const FROZEN_MESSAGE_PREFIX = "streaming-message-";

function textOf(raw: string[] | undefined): string {
  return (raw || []).join("");
}

function buildFrozenStreamingItems(current: ChatState, beforeItemId: string) {
  const baseId = `${current.activeTurnId || "turn"}-${current.activeStreamingRound ?? "round"}-${beforeItemId}`;
  const timestamp = Date.now();
  const items = [];

  if (current.streamingReasoning) {
    items.push({
      type: "reasoning" as const,
      id: `${FROZEN_REASONING_PREFIX}${baseId}`,
      summaryText: [],
      rawContent: [current.streamingReasoning],
      timestamp,
    });
  }

  if (current.streamingContent) {
    items.push({
      type: "assistant_message" as const,
      id: `${FROZEN_MESSAGE_PREFIX}${baseId}`,
      content: current.streamingContent,
      phase: "commentary" as const,
      timestamp,
    });
  }

  return items;
}

function findFrozenStreamingIndex(
  messages: ChatState["messages"],
  item: ChatState["messages"][number]
): number {
  if (item.type === "reasoning") {
    const completedText = textOf(item.rawContent);
    return messages.findIndex(
      (message) =>
        message.type === "reasoning" &&
        message.id.startsWith(FROZEN_REASONING_PREFIX) &&
        textOf(message.rawContent) === completedText
    );
  }

  if (item.type === "assistant_message") {
    return messages.findIndex(
      (message) =>
        message.type === "assistant_message" &&
        message.id.startsWith(FROZEN_MESSAGE_PREFIX) &&
        message.content === item.content
    );
  }

  return -1;
}

/**
 * 处理 Agent 事件，返回需要应用到 store 的状态 patches。
 * 
 * 参考 Codex 事件驱动模型：
 * - 所有消息只从 item_completed 事件产出
 * - turn_completed 只清理流式状态，不创建消息
 * - item_started 用于工具调用等需要即时展示的条目
 */
export function handleAgentEvent(
  event: AgentEvent,
  current: ChatState,
  patches: Array<Partial<ChatState>>
): Array<Partial<ChatState>> {
  if (event.threadId) {
    if (event.type === "turn_started") {
      const { [event.threadId]: _stoppedThread, ...stoppedRest } = current.stoppedThreadIds;
      patches.push({
        runningThreadIds: {
          ...current.runningThreadIds,
          [event.threadId]: true,
        },
        stoppedThreadIds: stoppedRest,
      });
    } else if (
      event.type === "turn_completed" ||
      event.type === "turn_interrupted" ||
      event.type === "turn_failed"
    ) {
      const { [event.threadId]: _completedThread, ...rest } = current.runningThreadIds;
      patches.push({ runningThreadIds: rest });
    }
  }

  const matchesActiveThread = Boolean(
    event.threadId && current.activeThreadId && event.threadId === current.activeThreadId
  );
  const matchesPendingClient = Boolean(
    event.clientId && current.activeClientId && event.clientId === current.activeClientId
  );

  if (event.threadId) {
    if (current.activeThreadId && !matchesActiveThread) {
      return patches;
    }
    if (!current.activeThreadId && !matchesPendingClient) {
      return patches;
    }
  }

  switch (event.type) {
    case "turn_started":
      patches.push({
        activeThreadId: !current.activeThreadId && matchesPendingClient && event.threadId
          ? event.threadId
          : current.activeThreadId,
        activeTurnId: event.turnId,
        turnStatus: "in_progress",
        isStreaming: true,
        streamingContent: "",
        streamingReasoning: "",
        activeStreamingRound: null,
        error: null,
      });
      break;

    case "turn_completed":
      patches.push({
        isStreaming: false,
        streamingContent: "",
        streamingReasoning: "",
        activeStreamingRound: null,
        activeTurnId: null,
        activeClientId: null,
        turnStatus: "completed",
        tokenUsage: event.usage || current.tokenUsage,
      });
      // Note: caller should call loadThreads() after applying patches
      break;

    case "turn_interrupted":
      patches.push({
        isStreaming: false,
        streamingContent: "",
        streamingReasoning: "",
        activeStreamingRound: null,
        activeTurnId: null,
        activeClientId: null,
        turnStatus: "interrupted",
        lastInterruptContext: "对话已被中断，你可以继续提问让 AI 从断点恢复",
      });
      break;

    case "turn_failed":
      patches.push({
        isStreaming: false,
        streamingContent: "",
        streamingReasoning: "",
        activeStreamingRound: null,
        activeTurnId: null,
        activeClientId: null,
        turnStatus: "failed",
        error: event.error,
      });
      break;

    case "item_started":
      if (event.item.type === "tool_call") {
        patches.push((() => {
          const exists = current.messages.some((m) => m.id === event.item.id);
          // tool_call 到达时立即清空流式缓冲，让 UI 从"思考态"切到"工具态"
          // 避免 reasoning streamDelta 和 item_completed 在两个 IPC 通道
          // 不同步导致 UI 卡在思考态
          const patch: Partial<ChatState> = {};
          if (!exists) {
            patch.messages = [
              ...current.messages,
              ...buildFrozenStreamingItems(current, event.item.id),
              event.item,
            ];
          }
          // 只有确实有思考过才清，避免纯文本情形被误清
          if (current.streamingReasoning) patch.streamingReasoning = "";
          if (current.streamingContent) patch.streamingContent = "";
          return patch;
        })());
      }
      if (event.item.type === "compact_progress") {
        const exists = current.messages.some((m) => m.id === event.item.id);
        if (!exists) {
          patches.push({ messages: [...current.messages, event.item] });
        }
      }
      break;

    case "item_completed": {
      const idx = current.messages.findIndex((m) => m.id === event.item.id);
      const newMessages = [...current.messages];
      if (idx >= 0) {
        newMessages[idx] = event.item;
      } else {
        const frozenIndex = findFrozenStreamingIndex(newMessages, event.item);
        if (frozenIndex >= 0) {
          newMessages[frozenIndex] = event.item;
        } else {
          newMessages.push(event.item);
        }
      }
      const patch: Partial<ChatState> = { messages: newMessages };
      if (event.item.type === "reasoning") {
        const completedReasoning = textOf(event.item.rawContent) || textOf(event.item.summaryText);
        if (current.streamingReasoning === completedReasoning) {
          patch.streamingReasoning = "";
        }
      }
      if (event.item.type === "assistant_message") {
        if (event.item.phase === "final") {
          patch.streamingContent = "";
          patch.streamingReasoning = "";
          patch.activeStreamingRound = null;
        } else if (current.streamingContent === event.item.content) {
          patch.streamingContent = "";
        }
      }
      patches.push(patch);
      break;
    }

    case "item_updated": {
      const idx = current.messages.findIndex((m) => m.id === event.item.id);
      if (idx >= 0) {
        const newMessages = [...current.messages];
        newMessages[idx] = event.item;
        patches.push({ messages: newMessages });
      }
      break;
    }

    case "thread_compact_started":
      patches.push({
        compactionNotice: `正在压缩上下文：${event.params.tokensBefore} tokens，失败最多重试 ${event.params.retryCount} 次`,
      });
      break;

    case "tool_approval_required":
      patches.push({
        pendingToolCall: {
          id: (event as any).toolCallId,
          toolName: (event as any).toolName,
          arguments: (event as any).arguments,
          riskLevel: (event as any).riskLevel,
          description: (event as any).description,
          sandboxJustification: (event as any).sandboxJustification,
        },
      });
      break;

    case "context_compacted":
      patches.push({
        compactionNotice: `上下文已压缩：${event.tokensBefore} → ${event.tokensAfter} tokens`,
      });
      break;

    case "context_usage":
      patches.push({
        contextUsage: {
          estimatedTokens: event.estimatedTokens,
          threshold: event.threshold,
          percentage: event.percentage,
          contextWindowSize: event.contextWindowSize,
        },
      });
      break;

    case "error":
      patches.push({
        isStreaming: false,
        error: event.message,
        turnStatus: "failed",
      });
      break;

    case "warning":
      patches.push({
        compactionNotice: event.message,
      });
      break;
  }

  return patches;
}
