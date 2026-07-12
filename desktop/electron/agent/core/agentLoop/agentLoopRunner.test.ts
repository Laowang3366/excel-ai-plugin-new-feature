import { describe, expect, it, vi } from "vitest";

import type { AIStreamEvent } from "../../providers/aiClient";
import type {
  AgentTurnCallbacks,
  Thread,
  Turn,
  TurnItem,
} from "../../shared/types";
import { runAgentLoopRounds } from "./agentLoopRunner";

async function* streamEvents(events: AIStreamEvent[]): AsyncIterable<AIStreamEvent> {
  for (const event of events) {
    yield event;
  }
}

function createTurn(): Turn {
  return {
    turnId: "turn-1",
    threadId: "thread-1",
    status: "in_progress",
    startedAt: 1,
    items: [],
  };
}

function createThread(): Thread {
  return {
    metadata: {
      threadId: "thread-1",
      preview: "",
      modelProvider: "test",
      model: "test-model",
      createdAt: 1,
      updatedAt: 1,
    },
    turns: [],
  };
}

describe("runAgentLoopRounds", () => {
  it("emits final assistant item and applies usage for a plain response", async () => {
    const usage = { inputTokens: 12, outputTokens: 3 };
    const turn = createTurn();
    const thread = createThread();
    const events: string[] = [];
    const aiClient = {
      streamChat: vi.fn(() => {
        expect(events).toEqual(["context_usage:none"]);
        return streamEvents([
          { type: "text_delta", delta: "完成" },
          { type: "usage", usage },
          { type: "done", finishReason: "stop" },
        ]);
      }),
    };
    const callbacks: AgentTurnCallbacks = {
      onStreamDelta: (delta, itemType, roundId) => {
        events.push(`delta:${itemType}:${roundId}:${delta}`);
      },
      onEvent: (event) => {
        const itemType = "item" in event ? event.item.type : "none";
        events.push(`${event.type}:${itemType}`);
      },
    };
    const appended: TurnItem[] = [];

    await runAgentLoopRounds({
      turn,
      callbacks,
      turnInput: { content: "测试" },
      aiClient: aiClient as any,
      aiConfig: {
        provider: "test",
        apiKey: "",
        baseUrl: "",
        model: "test-model",
      },
      toolExecutors: new Map(),
      approvalConfig: { permissionMode: "normal" },
      appendTurnItem: vi.fn(async (_threadId, _turnId, item) => {
        appended.push(item);
      }),
      getTurnItemGroups: () => [[{
        type: "user_message",
        id: "user-1",
        content: "测试",
        timestamp: 1,
      }]],
      getActiveThread: () => thread,
      getSessionCompactionConfig: () => ({
        enabled: false,
        contextWindowSize: 128_000,
        autoCompactTokenThreshold: 100_000,
        retainedUserMessageMaxTokens: 1000,
      }),
      runMidTurnCompaction: vi.fn(),
      emitContextUsage: (targetCallbacks) => {
        targetCallbacks.onEvent({
          type: "context_usage",
          estimatedTokens: 1,
          threshold: 100,
          percentage: 1,
          contextWindowSize: 128_000,
        });
      },
      throwIfAborted: vi.fn(),
    });

    expect(aiClient.streamChat).toHaveBeenCalledWith(expect.objectContaining({
      roundId: 1,
      reasoningMode: "high",
    }));
    expect(appended).toHaveLength(1);
    expect(appended[0]).toMatchObject({
      type: "assistant_message",
      phase: "final",
      content: "完成",
    });
    expect(turn.tokenUsage).toEqual(usage);
    expect(thread.metadata.totalTokenUsage).toEqual(usage);
    expect(events).toEqual([
      "context_usage:none",
      "delta:assistant_message:1:完成",
      "item_started:assistant_message",
      "item_completed:assistant_message",
    ]);
  });
});
