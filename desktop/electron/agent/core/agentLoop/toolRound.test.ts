import { describe, expect, it, vi } from "vitest";

import type { ToolCallItem, ToolDefinition, Turn, TurnItem } from "../../shared/types";
import type { StreamResult } from "./streamCollector";
import {
  handleToolRound,
  shouldRunMidTurnCompaction,
} from "./toolRound";
import { getLatestFormulaVerification } from "./formulaVerification";

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

  it("loads mandatory methodology and optional scene knowledge after formula.prepare", async () => {
    const prepareItem: ToolCallItem = {
      type: "tool_call",
      id: "prepare-1",
      toolName: "formula.prepare",
      arguments: {},
      status: "pending",
      timestamp: 1,
    };
    const processToolCallsImpl = vi.fn(async (
      toolCalls: Array<{ id: string; name: string; arguments: string }>,
      _pendingItems: unknown,
      turn: Turn,
    ) => {
      const toolCall = toolCalls[0];
      const args = JSON.parse(toolCall.arguments || "{}");
      if (!turn.items.some((item) => item.type === "tool_call" && item.id === toolCall.id)) {
        turn.items.push({
          type: "tool_call",
          id: toolCall.id,
          toolName: toolCall.name,
          arguments: args,
          status: "completed",
          timestamp: 2,
        });
      }
      turn.items.push({
        type: "tool_result",
        id: `result-${toolCall.id}`,
        toolCallId: toolCall.id,
        toolName: toolCall.name,
        result: toolCall.name === "formula.prepare"
          ? {
              status: "ready",
              scenario: "分组聚合",
              inputShape: "记录表",
              outputShape: "两列汇总表",
              inputGrain: "明细",
              outputGrain: "部门",
              businessKeys: ["部门"],
              transformChain: ["分组", "聚合"],
              constraints: [],
              acceptanceChecks: [{ type: "unique_key", description: "部门唯一", required: true, params: {} }],
              assumptions: [],
            }
          : args.scope === "formula_scene"
            ? { status: "no_match", matchCount: 0 }
            : "已返回公式解题方法论",
        isError: false,
        timestamp: 2,
      });
    });
    const turn = createTurn([
      userItem("【功能模块：生成公式】\n任务说明：按部门统计剩余餐费\n数据源选区：Sheet1!B3:E12\n答案填入锚点/选区：Sheet1!G3"),
      {
        type: "tool_call",
        id: "read-1",
        toolName: "range.read",
        arguments: { sheetName: "Sheet1", range: "B3:E12" },
        status: "completed",
        timestamp: 1,
      },
      {
        type: "tool_result",
        id: "result-read-1",
        toolCallId: "read-1",
        toolName: "range.read",
        result: [["部门", "人数", "餐时", "用餐人数"]],
        isError: false,
        timestamp: 2,
      },
      prepareItem,
    ]);

    const handled = await handleToolRound({
      streamResult: {
        assistantContent: "",
        reasoningContent: [],
        reasoningSummary: [],
        toolCalls: [{ id: "prepare-1", name: "formula.prepare", arguments: "{}" }],
        finishReason: "tool_calls",
        usage: undefined,
        pendingToolCallItems: new Map([["prepare-1", prepareItem]]),
      },
      turn,
      toolExecutors: new Map([
        ["range.read", { name: "range.read", execute: vi.fn() }],
        ["formula.prepare", { name: "formula.prepare", execute: vi.fn() }],
        ["formula.verify", { name: "formula.verify", execute: vi.fn() }],
        ["knowledge.search", { name: "knowledge.search", execute: vi.fn() }],
      ]),
      approvalConfig: { permissionMode: "normal" },
      callbacks: { onEvent: vi.fn() },
      appendTurnItem: vi.fn(),
      turnItemGroups: [[userItem("【功能模块：生成公式】")]],
      effectiveSystemPrompt: "system",
      toolDefs: [],
      compactionConfig: { ...compactionConfig, enabled: false },
      runMidTurnCompaction: vi.fn(),
      throwIfAborted: vi.fn(),
      processToolCallsImpl: processToolCallsImpl as any,
    });

    expect(handled).toBe(true);
    expect(processToolCallsImpl).toHaveBeenCalledTimes(3);
    const methodologyCall = processToolCallsImpl.mock.calls[1][0][0];
    const sceneCall = processToolCallsImpl.mock.calls[2][0][0];
    expect(methodologyCall.name).toBe("knowledge.search");
    expect(JSON.parse(methodologyCall.arguments)).toMatchObject({
      topK: 3,
      scope: "formula_methodology",
    });
    expect(JSON.parse(methodologyCall.arguments).query).toContain("场景：分组聚合");
    expect(JSON.parse(sceneCall.arguments)).toMatchObject({ topK: 3, scope: "formula_scene" });
  });

  it("automatically reads and verifies a successful formula write", async () => {
    const writeItem: ToolCallItem = {
      type: "tool_call",
      id: "write-1",
      toolName: "range.write",
      arguments: { sheetName: "Sheet1", range: "G1", values: [["=SUM(A2:A3)"]] },
      status: "pending",
      timestamp: 10,
    };
    const processToolCallsImpl = vi.fn(async (
      toolCalls: Array<{ id: string; name: string; arguments: string }>,
      _pendingItems: unknown,
      turn: Turn,
    ) => {
      const toolCall = toolCalls[0];
      const args = JSON.parse(toolCall.arguments || "{}");
      if (!turn.items.some((item) => item.type === "tool_call" && item.id === toolCall.id)) {
        turn.items.push({ type: "tool_call", id: toolCall.id, toolName: toolCall.name, arguments: args, status: "completed", timestamp: 11 });
      }
      turn.items.push({
        type: "tool_result",
        id: `result-${toolCall.id}`,
        toolCallId: toolCall.id,
        toolName: toolCall.name,
        result: toolCall.name === "range.read"
          ? { address: "G1", values: [[30]], expanded: false }
          : toolCall.name === "formula.verify"
            ? { status: "passed", writeToolCallId: "write-1", checks: [] }
            : "写入成功",
        isError: false,
        timestamp: 12,
      });
    });
    const prepared = {
      status: "ready",
      scenario: "标量聚合",
      inputShape: "向量",
      outputShape: "单值",
      inputGrain: "明细",
      outputGrain: "合计",
      businessKeys: [],
      transformChain: ["求和"],
      constraints: [],
      acceptanceChecks: [{ type: "shape", description: "输出单值", required: true, params: { expectedRows: 1, expectedColumns: 1 } }],
      assumptions: [],
    };
    const turn = createTurn([
      userItem("【功能模块：生成公式】\n任务说明：计算合计\n数据源选区：Sheet1!A1:A3\n答案填入锚点/选区：Sheet1!G1"),
      { type: "tool_call", id: "read-source", toolName: "range.read", arguments: { sheetName: "Sheet1", range: "A1:A3" }, status: "completed", timestamp: 1 },
      { type: "tool_result", id: "result-read-source", toolCallId: "read-source", toolName: "range.read", result: [["金额"], [10], [20]], isError: false, timestamp: 2 },
      { type: "tool_call", id: "prepare", toolName: "formula.prepare", arguments: prepared, status: "completed", timestamp: 3 },
      { type: "tool_result", id: "result-prepare", toolCallId: "prepare", toolName: "formula.prepare", result: prepared, isError: false, timestamp: 4 },
      { type: "tool_call", id: "method", toolName: "knowledge.search", arguments: { query: "合计", scope: "formula_methodology" }, status: "completed", timestamp: 5 },
      { type: "tool_result", id: "result-method", toolCallId: "method", toolName: "knowledge.search", result: "方法论", isError: false, timestamp: 6 },
      { type: "tool_call", id: "scene", toolName: "knowledge.search", arguments: { query: "合计", scope: "formula_scene" }, status: "completed", timestamp: 7 },
      { type: "tool_result", id: "result-scene", toolCallId: "scene", toolName: "knowledge.search", result: { status: "no_match" }, isError: false, timestamp: 8 },
      writeItem,
    ]);

    await handleToolRound({
      streamResult: {
        assistantContent: "",
        reasoningContent: [],
        reasoningSummary: [],
        toolCalls: [{ id: "write-1", name: "range.write", arguments: JSON.stringify(writeItem.arguments) }],
        finishReason: "tool_calls",
        usage: undefined,
        pendingToolCallItems: new Map([["write-1", writeItem]]),
      },
      turn,
      toolExecutors: new Map([
        ["range.write", { name: "range.write", execute: vi.fn() }],
        ["range.read", { name: "range.read", execute: vi.fn() }],
        ["formula.prepare", { name: "formula.prepare", execute: vi.fn() }],
        ["formula.verify", { name: "formula.verify", execute: vi.fn() }],
        ["knowledge.search", { name: "knowledge.search", execute: vi.fn() }],
      ]),
      approvalConfig: { permissionMode: "confirm_all" },
      callbacks: { onEvent: vi.fn() },
      appendTurnItem: vi.fn(),
      turnItemGroups: [[userItem("【功能模块：生成公式】")]],
      effectiveSystemPrompt: "system",
      toolDefs: [],
      compactionConfig: { ...compactionConfig, enabled: false },
      runMidTurnCompaction: vi.fn(),
      throwIfAborted: vi.fn(),
      processToolCallsImpl: processToolCallsImpl as any,
    });

    expect(processToolCallsImpl).toHaveBeenCalledTimes(3);
    expect(processToolCallsImpl.mock.calls.map((entry) => entry[0][0].name)).toEqual([
      "range.write",
      "range.read",
      "formula.verify",
    ]);
    expect(JSON.parse(processToolCallsImpl.mock.calls[1][0][0].arguments)).toMatchObject({ range: "G1", expand: "spill" });
  });

  it("completes the workflow through real tool processing", async () => {
    const payload = "【功能模块：生成公式】\n任务说明：计算金额合计\n数据源选区：Sheet1!A1:A3\n答案填入锚点/选区：Sheet1!G1";
    const prepareArgs = {
      status: "ready",
      scenario: "标量聚合",
      inputShape: "带表头的数值向量",
      outputShape: "单值",
      inputGrain: "每行一个金额",
      outputGrain: "总金额",
      businessKeys: [],
      transformChain: ["忽略表头", "求和"],
      constraints: ["WPS"],
      acceptanceChecks: [{ type: "shape", description: "输出一个单元格", required: true, params: { expectedRows: 1, expectedColumns: 1 } }],
      assumptions: [],
    };
    const prepareItem: ToolCallItem = {
      type: "tool_call",
      id: "prepare-real",
      toolName: "formula.prepare",
      arguments: prepareArgs,
      status: "pending",
      timestamp: 3,
    };
    const turn = createTurn([
      userItem(payload),
      { type: "tool_call", id: "read-real", toolName: "range.read", arguments: { sheetName: "Sheet1", range: "A1:A3" }, status: "completed", timestamp: 1 },
      { type: "tool_result", id: "result-read-real", toolCallId: "read-real", toolName: "range.read", result: [["金额"], [10], [20]], isError: false, timestamp: 2 },
      prepareItem,
    ]);
    const executors = new Map([
      ["formula.prepare", {
        name: "formula.prepare",
        execute: vi.fn(async () => ({ success: true, data: prepareArgs })),
      }],
      ["formula.verify", {
        name: "formula.verify",
        execute: vi.fn(async () => ({ success: false, error: "runtime only" })),
      }],
      ["knowledge.search", {
        name: "knowledge.search",
        execute: vi.fn(async (args: Record<string, unknown>) => ({
          success: true,
          data: args.scope === "formula_scene" ? { status: "no_match", matchCount: 0 } : "方法论",
        })),
      }],
      ["range.write", {
        name: "range.write",
        execute: vi.fn(async () => ({ success: true, data: "写入成功" })),
      }],
      ["range.read", {
        name: "range.read",
        execute: vi.fn(async () => ({ success: true, data: { address: "G1", expanded: false, values: [[30]] } })),
      }],
    ]);
    const common = {
      turn,
      toolExecutors: executors,
      approvalConfig: { permissionMode: "confirm_all" as const },
      callbacks: { onEvent: vi.fn() },
      appendTurnItem: vi.fn(async () => {}),
      turnItemGroups: [turn.items],
      effectiveSystemPrompt: "system",
      toolDefs: [],
      compactionConfig: { ...compactionConfig, enabled: false },
      runMidTurnCompaction: vi.fn(async () => {}),
      throwIfAborted: vi.fn(),
    };

    await handleToolRound({
      ...common,
      streamResult: {
        assistantContent: "",
        reasoningContent: [],
        reasoningSummary: [],
        toolCalls: [{ id: "prepare-real", name: "formula.prepare", arguments: JSON.stringify(prepareArgs) }],
        finishReason: "tool_calls",
        usage: undefined,
        pendingToolCallItems: new Map([["prepare-real", prepareItem]]),
      },
    });
    expect(turn.items.filter((item) => item.type === "tool_result" && item.toolName === "knowledge.search")).toHaveLength(2);

    const writeItem: ToolCallItem = {
      type: "tool_call",
      id: "write-real",
      toolName: "range.write",
      arguments: { sheetName: "Sheet1", range: "G1", values: [["=SUM(A2:A3)"]] },
      status: "pending",
      timestamp: 10,
    };
    turn.items.push(writeItem);
    await handleToolRound({
      ...common,
      turnItemGroups: [turn.items],
      streamResult: {
        assistantContent: "",
        reasoningContent: [],
        reasoningSummary: [],
        toolCalls: [{ id: "write-real", name: "range.write", arguments: JSON.stringify(writeItem.arguments) }],
        finishReason: "tool_calls",
        usage: undefined,
        pendingToolCallItems: new Map([["write-real", writeItem]]),
      },
    });

    expect(getLatestFormulaVerification(turn)).toMatchObject({
      status: "passed",
      writeToolCallId: "write-real",
      actualShape: { rows: 1, columns: 1 },
    });
  });

  it("calculates mid-turn compaction from prompt tokens", () => {
    const toolDef: ToolDefinition = {
      name: "test.tool",
      description: "测试工具",
      parameters: {},
      riskLevel: "safe",
      requiresApproval: false,
    };

    expect(shouldRunMidTurnCompaction({
      turnItemGroups: [[userItem("这是一段足够长的内容，用于触发 token 估算超过阈值。")]],
      systemPrompt: "system",
      tools: [toolDef],
      compactionConfig,
    })).toBe(true);
  });
});
