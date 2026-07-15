import { describe, expect, it, vi } from "vitest";

import type { AgentTurnCallbacks, ToolCallItem, Turn } from "../../shared/types";
import type { StreamResult } from "./streamCollector";
import { emitStreamResultItems } from "./streamResultItems";

function createTurn(): Turn {
  return {
    turnId: "turn-1",
    threadId: "thread-1",
    status: "in_progress",
    startedAt: 1,
    items: [],
  };
}

function createCallbacks(events: unknown[]): AgentTurnCallbacks {
  return {
    onEvent: (event) => events.push(event),
  };
}

describe("emitStreamResultItems", () => {
  it("emits reasoning, assistant message, then tool call in stream order", async () => {
    const toolItem: ToolCallItem = {
      type: "tool_call",
      id: "tool-1",
      toolName: "range.read",
      arguments: { range: "A1" },
      status: "pending",
      timestamp: 1,
    };
    const streamResult: StreamResult = {
      assistantContent: "我先读取数据。",
      reasoningContent: ["分析中"],
      reasoningSummary: ["需要读取表格"],
      toolCalls: [{ id: "tool-1", name: "range.read", arguments: "{}" }],
      finishReason: "tool_calls",
      usage: undefined,
      pendingToolCallItems: new Map([["tool-1", toolItem]]),
    };
    const turn = createTurn();
    const events: unknown[] = [];
    const appendTurnItem = vi.fn().mockResolvedValue(undefined);

    await emitStreamResultItems({
      streamResult,
      turn,
      callbacks: createCallbacks(events),
      appendTurnItem,
    });

    expect(turn.items.map((item) => item.type)).toEqual([
      "reasoning",
      "assistant_message",
      "tool_call",
    ]);
    expect(appendTurnItem.mock.calls.map((call) => call[2].type)).toEqual([
      "reasoning",
      "assistant_message",
      "tool_call",
    ]);
    expect(
      events.map(
        (event) => (event as { type: string; item?: { type: string } }).item?.type ?? "none",
      ),
    ).toEqual(["reasoning", "reasoning", "assistant_message", "assistant_message", "tool_call"]);
    expect(turn.items[1]).toMatchObject({
      type: "assistant_message",
      phase: "commentary",
      content: "我先读取数据。",
    });
  });

  it("marks assistant message as final when there are no tool calls", async () => {
    const streamResult: StreamResult = {
      assistantContent: "完成。",
      reasoningContent: [],
      reasoningSummary: [],
      toolCalls: [],
      finishReason: "stop",
      usage: undefined,
      pendingToolCallItems: new Map(),
    };
    const turn = createTurn();

    await emitStreamResultItems({
      streamResult,
      turn,
      callbacks: createCallbacks([]),
      appendTurnItem: vi.fn().mockResolvedValue(undefined),
    });

    expect(turn.items).toHaveLength(1);
    expect(turn.items[0]).toMatchObject({
      type: "assistant_message",
      phase: "final",
    });
  });
});
