import { describe, expect, it, vi } from "vitest";

import type { ToolCallItem, ToolDefinition, Turn, TurnItem } from "../../shared/types";
import type { StreamResult } from "./streamCollector";
import { handleToolRound, shouldRunMidTurnCompaction } from "./toolRound";

function createTurn(items: TurnItem[] = []): Turn {
  return {
    threadId: "thread-1",
    turnId: "turn-1",
    status: "in_progress",
    startedAt: 1,
    items,
  };
}

const userItem = (content: string): TurnItem => ({
  type: "user_message",
  id: "u1",
  content,
  timestamp: 1,
});

function streamResultWithTool(): StreamResult {
  const toolItem: ToolCallItem = {
    type: "tool_call",
    id: "call-1",
    toolName: "test.tool",
    arguments: {},
    status: "pending",
    timestamp: 1,
  };
  return {
    assistantContent: "",
    reasoningContent: [],
    reasoningSummary: [],
    toolCalls: [{ id: "call-1", name: "test.tool", arguments: "{}" }],
    finishReason: "tool_calls",
    usage: undefined,
    pendingToolCallItems: new Map([["call-1", toolItem]]),
  };
}

const compactionConfig = {
  enabled: true,
  contextWindowSize: 1000,
  autoCompactTokenThreshold: 1,
  retainedUserMessageMaxTokens: 100,
};

describe("toolRound", () => {
  it("skips rounds without tool calls", async () => {
    const handled = await handleToolRound({
      streamResult: {
        assistantContent: "done",
        reasoningContent: [],
        reasoningSummary: [],
        toolCalls: [],
        finishReason: "stop",
        usage: undefined,
        pendingToolCallItems: new Map(),
      },
      turn: createTurn(),
      toolExecutors: new Map(),
      approvalConfig: { permissionMode: "confirm_all" },
      callbacks: { onEvent: vi.fn() },
      appendTurnItem: vi.fn(),
      turnItemGroups: [],
      effectiveSystemPrompt: "",
      toolDefs: [],
      compactionConfig,
      runMidTurnCompaction: vi.fn(),
      throwIfAborted: vi.fn(),
    });

    expect(handled).toBe(false);
  });

  it("processes tool calls and runs mid-turn compaction when over threshold", async () => {
    const processToolCallsImpl = vi.fn().mockResolvedValue(undefined);
    const runMidTurnCompaction = vi.fn().mockResolvedValue(undefined);
    const throwIfAborted = vi.fn();

    const handled = await handleToolRound({
      streamResult: streamResultWithTool(),
      turn: createTurn([userItem("请读取并继续分析")]),
      toolExecutors: new Map(),
      approvalConfig: { permissionMode: "confirm_all" },
      callbacks: { onEvent: vi.fn() },
      appendTurnItem: vi.fn(),
      appendToolExecutionLog: vi.fn(),
      turnItemGroups: [[userItem("请读取并继续分析")]],
      effectiveSystemPrompt: "system",
      toolDefs: [],
      compactionConfig,
      runMidTurnCompaction,
      throwIfAborted,
      processToolCallsImpl,
    });

    expect(handled).toBe(true);
    expect(processToolCallsImpl).toHaveBeenCalledOnce();
    expect(runMidTurnCompaction).toHaveBeenCalledOnce();
    expect(throwIfAborted).toHaveBeenCalledTimes(2);
  });

  it("calculates mid-turn compaction from prompt tokens", () => {
    const toolDef: ToolDefinition = {
      name: "test.tool",
      description: "测试工具",
      parameters: {},
      riskLevel: "safe",
      requiresApproval: false,
    };

    expect(
      shouldRunMidTurnCompaction({
        turnItemGroups: [[userItem("这是一段足够长的内容，用于触发 token 估算超过阈值。")]],
        systemPrompt: "system",
        tools: [toolDef],
        compactionConfig,
      }),
    ).toBe(true);
  });
});
