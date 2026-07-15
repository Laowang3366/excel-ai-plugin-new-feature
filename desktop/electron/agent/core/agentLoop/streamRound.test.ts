import { describe, expect, it, vi } from "vitest";

import type { AIStreamEvent } from "../../providers/aiClient";
import type { Thread, Turn, TurnItem } from "../../shared/types";
import { applyStreamUsage, collectRoundStream, emitStreamErrorItem } from "./streamRound";

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
      modelProvider: "openai",
      createdAt: 1,
      updatedAt: 1,
    },
    turns: [],
  };
}

describe("streamRound", () => {
  it("collects a model sampling stream with retry wrapper", async () => {
    const aiClient = {
      streamChat: vi.fn(() =>
        streamEvents([
          { type: "text_delta", delta: "你好" },
          { type: "done", finishReason: "stop" },
        ]),
      ),
    };
    const streamResult = await collectRoundStream({
      aiClient: aiClient as never,
      streamParams: {
        messages: [],
        tools: [],
        systemPrompt: "system",
        maxTokens: 100,
        reasoningMode: "high",
        roundId: 2,
      },
      callbacks: { onEvent: vi.fn(), onStreamDelta: vi.fn() },
      round: 2,
    });

    expect(streamResult.assistantContent).toBe("你好");
    expect(aiClient.streamChat).toHaveBeenCalledTimes(1);
  });

  it("retries transient failures before any output becomes visible", async () => {
    const aiClient = {
      streamChat: vi
        .fn()
        .mockImplementationOnce(() =>
          streamEvents([{ type: "error", error: "API 请求失败 (503): unavailable" }]),
        )
        .mockImplementationOnce(() =>
          streamEvents([
            { type: "text_delta", delta: "重试成功" },
            { type: "done", finishReason: "stop" },
          ]),
        ),
    };
    const onStreamDelta = vi.fn();

    const result = await collectRoundStream({
      aiClient: aiClient as never,
      streamParams: {
        messages: [],
        tools: [],
        systemPrompt: "system",
        maxTokens: 100,
        reasoningMode: "off",
      },
      callbacks: { onEvent: vi.fn(), onStreamDelta },
      round: 1,
      retryConfig: { maxRetries: 1, baseDelayMs: 0, maxDelayMs: 0 },
    });

    expect(result.assistantContent).toBe("重试成功");
    expect(aiClient.streamChat).toHaveBeenCalledTimes(2);
    expect(onStreamDelta).toHaveBeenCalledTimes(1);
  });

  it("does not retry after a text delta has been emitted", async () => {
    const aiClient = {
      streamChat: vi.fn(() =>
        streamEvents([
          { type: "text_delta", delta: "部分内容" },
          { type: "error", error: "API 请求失败 (503): disconnected" },
        ]),
      ),
    };
    const onStreamDelta = vi.fn();

    await expect(
      collectRoundStream({
        aiClient: aiClient as never,
        streamParams: {
          messages: [],
          tools: [],
          systemPrompt: "system",
          maxTokens: 100,
          reasoningMode: "off",
        },
        callbacks: { onEvent: vi.fn(), onStreamDelta },
        round: 1,
        retryConfig: { maxRetries: 1, baseDelayMs: 0, maxDelayMs: 0 },
      }),
    ).rejects.toThrow("disconnected");

    expect(aiClient.streamChat).toHaveBeenCalledTimes(1);
    expect(onStreamDelta).toHaveBeenCalledTimes(1);
    expect(onStreamDelta).toHaveBeenCalledWith("部分内容", "assistant_message", 1);
  });

  it("does not retry after a tool item has been emitted", async () => {
    const aiClient = {
      streamChat: vi.fn(() =>
        streamEvents([
          { type: "tool_call_begin", toolCallId: "call-1", toolName: "range.read" },
          { type: "error", error: "API 请求失败 (503): disconnected" },
        ]),
      ),
    };
    const onEvent = vi.fn();

    await expect(
      collectRoundStream({
        aiClient: aiClient as never,
        streamParams: {
          messages: [],
          tools: [],
          systemPrompt: "system",
          maxTokens: 100,
          reasoningMode: "off",
        },
        callbacks: { onEvent, onStreamDelta: vi.fn() },
        round: 1,
        retryConfig: { maxRetries: 1, baseDelayMs: 0, maxDelayMs: 0 },
      }),
    ).rejects.toThrow("disconnected");

    expect(aiClient.streamChat).toHaveBeenCalledTimes(1);
    expect(onEvent).toHaveBeenCalledTimes(1);
    expect(onEvent).toHaveBeenCalledWith(
      expect.objectContaining({
        type: "item_started",
        item: expect.objectContaining({ id: "call-1", type: "tool_call" }),
      }),
    );
  });

  it("emits stream error items and stops the round", async () => {
    const turn = createTurn();
    const events: unknown[] = [];
    const errorItem: TurnItem = {
      type: "error",
      id: "error-1",
      message: "失败",
      timestamp: 1,
    };
    const appendTurnItem = vi.fn().mockResolvedValue(undefined);

    const handled = await emitStreamErrorItem({
      streamResult: {
        assistantContent: "",
        reasoningContent: [],
        reasoningSummary: [],
        toolCalls: [],
        finishReason: "error",
        usage: undefined,
        pendingToolCallItems: new Map(),
        errorItem,
      } as never,
      turn,
      callbacks: { onEvent: (event) => events.push(event) },
      appendTurnItem,
    });

    expect(handled).toBe(true);
    expect(turn.items).toEqual([errorItem]);
    expect(appendTurnItem).toHaveBeenCalledWith("thread-1", "turn-1", errorItem);
    expect(events.map((event) => (event as { type: string }).type)).toEqual([
      "item_started",
      "item_completed",
      "error",
    ]);
  });

  it("applies token usage to the turn and active thread", () => {
    const turn = createTurn();
    const thread = createThread();

    applyStreamUsage({
      streamResult: {
        assistantContent: "",
        reasoningContent: [],
        reasoningSummary: [],
        toolCalls: [],
        finishReason: "stop",
        usage: { inputTokens: 2, outputTokens: 3 },
        pendingToolCallItems: new Map(),
      },
      turn,
      activeThread: thread,
    });

    expect(turn.tokenUsage).toEqual({ inputTokens: 2, outputTokens: 3 });
    expect(thread.metadata.totalTokenUsage).toEqual({ inputTokens: 2, outputTokens: 3 });
  });
});
