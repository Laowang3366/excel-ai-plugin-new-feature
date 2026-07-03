import { describe, expect, it, vi } from "vitest";

import type { AIStreamEvent } from "../../providers/aiClient";
import type { AgentTurnCallbacks, TurnItem } from "../../shared/types";
import { collectStreamEvents, emitInterruptedProgress, type StreamResult } from "./streamCollector";

async function* streamEvents(events: AIStreamEvent[]): AsyncIterable<AIStreamEvent> {
  for (const event of events) {
    yield event;
  }
}

function createCallbacks(): AgentTurnCallbacks {
  return {
    onEvent: vi.fn(),
    onStreamDelta: vi.fn(),
  };
}

describe("collectStreamEvents", () => {
  it("throws retriable API error events so the caller can retry the request", async () => {
    await expect(
      collectStreamEvents(
        streamEvents([
          { type: "error", error: "API 请求失败 (429): rate limit" },
        ]),
        createCallbacks()
      )
    ).rejects.toMatchObject({
      message: "API 请求失败 (429): rate limit",
    });
  });

  it("collects text, reasoning, usage and completed tool calls", async () => {
    const callbacks = createCallbacks();

    const result = await collectStreamEvents(
      streamEvents([
        { type: "text_delta", delta: "你好" },
        { type: "reasoning_delta", delta: "检查参数" },
        { type: "reasoning_summary_delta", delta: "已检查" },
        { type: "tool_call_begin", toolCallId: "call-1", toolName: "range.read" },
        { type: "tool_call_delta", toolCallId: "call-1", delta: "{\"sheetName\":" },
        { type: "tool_call_delta", toolCallId: "call-1", delta: "\"Sheet1\"}" },
        {
          type: "tool_call_end",
          toolCallId: "call-1",
          toolName: "range.read",
          arguments: "{\"sheetName\":\"Sheet1\"}",
        },
        { type: "usage", usage: { inputTokens: 10, outputTokens: 3 } },
        { type: "done", finishReason: "tool_calls" },
      ]),
      callbacks,
      2
    );

    expect(result.assistantContent).toBe("你好");
    expect(result.reasoningContent).toEqual(["检查参数"]);
    expect(result.reasoningSummary).toEqual(["已检查"]);
    expect(result.toolCalls).toEqual([
      { id: "call-1", name: "range.read", arguments: "{\"sheetName\":\"Sheet1\"}" },
    ]);
    expect(result.usage).toEqual({ inputTokens: 10, outputTokens: 3 });
    expect(result.finishReason).toBe("tool_calls");
    expect(result.pendingToolCallItems.get("call-1")).toMatchObject({
      type: "tool_call",
      id: "call-1",
      toolName: "range.read",
      arguments: { sheetName: "Sheet1" },
      status: "pending",
    });
    expect(callbacks.onStreamDelta).toHaveBeenCalledWith("你好", "assistant_message", 2);
    expect(callbacks.onStreamDelta).toHaveBeenCalledWith("检查参数", "reasoning", 2);
    expect(callbacks.onEvent).toHaveBeenCalledWith(expect.objectContaining({ type: "item_started" }));
    expect(callbacks.onEvent).toHaveBeenCalledWith(expect.objectContaining({ type: "item_updated" }));
  });

  it("keeps raw tool arguments when JSON parsing fails", async () => {
    const callbacks = createCallbacks();

    const result = await collectStreamEvents(
      streamEvents([
        { type: "tool_call_begin", toolCallId: "call-raw", toolName: "word.insertText" },
        {
          type: "tool_call_end",
          toolCallId: "call-raw",
          toolName: "word.insertText",
          arguments: "{bad-json",
        },
      ]),
      callbacks
    );

    expect(result.pendingToolCallItems.get("call-raw")?.arguments).toEqual({ _raw: "{bad-json" });
    expect(result.toolCalls).toEqual([
      { id: "call-raw", name: "word.insertText", arguments: "{bad-json" },
    ]);
  });

  it("returns an error item when the stream emits an error event", async () => {
    const callbacks = createCallbacks();

    const result = await collectStreamEvents(
      streamEvents([
        { type: "text_delta", delta: "partial" },
        { type: "error", error: "模型流中断" },
      ]),
      callbacks
    );

    expect(result.assistantContent).toBe("");
    expect((result as StreamResult & { errorItem?: TurnItem }).errorItem).toMatchObject({
      type: "error",
      message: "模型流中断",
    });
  });
});

describe("emitInterruptedProgress", () => {
  it("emits collected assistant and reasoning items when a turn is interrupted", async () => {
    const callbacks = createCallbacks();
    const turnItems: TurnItem[] = [];

    await emitInterruptedProgress(
      {
        assistantContent: "已经完成一半",
        reasoningContent: ["读取文件", "准备写入"],
        reasoningSummary: ["已读取"],
        toolCalls: [],
        finishReason: "",
        usage: undefined,
        pendingToolCallItems: new Map(),
      },
      turnItems,
      "thread-1",
      "turn-1",
      callbacks
    );

    expect(turnItems).toHaveLength(2);
    expect(turnItems[0]).toMatchObject({
      type: "assistant_message",
      content: "已经完成一半",
      phase: "commentary",
    });
    expect(turnItems[1]).toMatchObject({
      type: "reasoning",
      summaryText: ["已读取"],
      rawContent: ["读取文件", "准备写入"],
    });
    expect(callbacks.onEvent).toHaveBeenCalledTimes(4);
  });
});
