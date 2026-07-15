import { beforeEach, describe, expect, it, vi } from "vitest";

import type {
  AgentTurnCallbacks,
  ToolCallItem,
  ToolExecutor,
  Turn,
  TurnItem,
} from "../../shared/types";
import type { ToolCallInfo } from "./streamCollector";
import {
  clearAlwaysAllowedTools,
  executeTool,
  getAlwaysAllowedTools,
  processToolCalls,
  shouldRequireApproval,
} from "./toolExecutor";
import { summarizeForLog } from "./toolExecutionLog";

function createTurn(item?: ToolCallItem): Turn {
  return {
    turnId: "turn-1",
    threadId: "thread-1",
    status: "in_progress",
    items: item ? [item] : [],
    startedAt: 1000,
  };
}

function createCallbacks(): AgentTurnCallbacks {
  return { onEvent: vi.fn() };
}

describe("executeTool", () => {
  it("returns a clear error for unknown tools", async () => {
    await expect(executeTool("missing.tool", "{}", new Map())).resolves.toEqual({
      success: false,
      error: "未知工具: missing.tool",
    });
  });

  it("parses JSON arguments before calling the executor", async () => {
    const executor = vi.fn(async () => ({ success: true, data: "ok" }));
    const executors = new Map<string, ToolExecutor>([
      ["range.read", { name: "range.read", execute: executor }],
    ]);
    const result = await executeTool(
      "range.read",
      "{\"sheetName\":\"Sheet1\",\"range\":\"A1\"}",
      executors
    );

    expect(result).toEqual({ success: true, data: "ok" });
    expect(executor).toHaveBeenCalledWith({ sheetName: "Sheet1", range: "A1" });
  });

  it("executes underscored tool aliases through the canonical executor when available", async () => {
    const executor = vi.fn(async () => ({ success: true, data: "parsed" }));
    const executors = new Map<string, ToolExecutor>([
      ["ocr.parseDocument", { name: "ocr.parseDocument", execute: executor }],
    ]);

    const result = await executeTool(
      "ocr_parseDocument",
      "{\"filePaths\":[\"C:\\\\Users\\\\29721\\\\Pictures\\\\image.png\"]}",
      executors
    );

    expect(result).toEqual({ success: true, data: "parsed" });
    expect(executor).toHaveBeenCalledWith({ filePaths: ["C:\\Users\\29721\\Pictures\\image.png"] });
  });

  it("rejects malformed and undeclared arguments before calling the executor", async () => {
    const execute = vi.fn(async () => ({ success: true }));
    const executors = new Map<string, ToolExecutor>([["range.read", { name: "range.read", execute }]]);

    await expect(executeTool("range.read", "[]", executors)).resolves.toMatchObject({
      success: false,
      error: expect.stringContaining("根节点必须为对象"),
    });
    await expect(executeTool(
      "range.read",
      '{"sheetName":"Sheet1","range":"A1","unexpected":true}',
      executors,
    )).resolves.toMatchObject({
      success: false,
      error: expect.stringContaining("未声明参数"),
    });
    expect(execute).not.toHaveBeenCalled();
  });
});

describe("shouldRequireApproval", () => {
  beforeEach(() => {
    clearAlwaysAllowedTools();
  });

  it("honors permission mode and always-allowed overrides", () => {
    expect(shouldRequireApproval("range.read", "normal")).toBe(true);
    expect(shouldRequireApproval("range.read", "auto_approve_safe")).toBe(false);
    expect(shouldRequireApproval("range.clear", "auto_approve_safe")).toBe(true);
    expect(shouldRequireApproval("range.clear", "confirm_all")).toBe(true);
    expect(shouldRequireApproval("macro.run", "confirm_all")).toBe(true);
    expect(shouldRequireApproval("unknown.tool", "confirm_all")).toBe(true);
  });
});

describe("processToolCalls", () => {
  beforeEach(() => {
    clearAlwaysAllowedTools();
  });

  it("does not execute a tool when approval is denied", async () => {
    const activeItem: ToolCallItem = {
      type: "tool_call",
      id: "call-1",
      toolName: "range.clear",
      arguments: { sheetName: "Sheet1", range: "A1" },
      status: "pending",
      timestamp: 1000,
    };
    const turn = createTurn(activeItem);
    const execute = vi.fn(async () => ({ success: true, data: "cleared" }));
    const callbacks = createCallbacks();
    const appended: TurnItem[] = [];

    await processToolCalls(
      [{ id: "call-1", name: "range.clear", arguments: "{\"sheetName\":\"Sheet1\",\"range\":\"A1\"}" }],
      new Map([["call-1", activeItem]]),
      turn,
      new Map([["range.clear", { name: "range.clear", execute }]]),
      {
        permissionMode: "normal",
        requestToolApproval: vi.fn(async () => ({ approved: false })),
      },
      callbacks,
      vi.fn(async (_threadId, _turnId, item) => {
        appended.push(item);
      })
    );

    expect(execute).not.toHaveBeenCalled();
    expect(activeItem.status).toBe("failed");
    expect(appended[appended.length - 1]).toMatchObject({
      type: "tool_result",
      toolCallId: "call-1",
      toolName: "range.clear",
      result: "用户取消了工具执行",
      isError: true,
    });
    expect(callbacks.onEvent).toHaveBeenCalledWith({ type: "item_updated", item: activeItem });
  });

  it("rejects invalid arguments before requesting approval", async () => {
    const execute = vi.fn(async () => ({ success: true }));
    const requestToolApproval = vi.fn(async () => ({ approved: true }));
    const turn = createTurn();

    await processToolCalls(
      [{ id: "call-invalid", name: "range.clear", arguments: '{"sheetName":"Sheet1","range":"A1","unknown":1}' }],
      new Map(),
      turn,
      new Map([["range.clear", { name: "range.clear", execute }]]),
      { permissionMode: "confirm_all", requestToolApproval },
      createCallbacks(),
      vi.fn(async () => {}),
    );

    expect(requestToolApproval).not.toHaveBeenCalled();
    expect(execute).not.toHaveBeenCalled();
    expect(turn.items).toEqual(expect.arrayContaining([
      expect.objectContaining({ type: "tool_call", status: "failed" }),
      expect.objectContaining({ type: "tool_result", isError: true, result: expect.stringContaining("未声明参数") }),
    ]));
  });

  it("creates a fallback tool_call item when the stream did not create one", async () => {
    const turn = createTurn();
    const execute = vi.fn(async () => ({ success: true, data: { values: [[1]] } }));
    const callbacks = createCallbacks();

    await processToolCalls(
      [{ id: "call-2", name: "range.read", arguments: "{\"sheetName\":\"Sheet1\",\"range\":\"A1\"}" }],
      new Map(),
      turn,
      new Map([["range.read", { name: "range.read", execute }]]),
      { permissionMode: "confirm_all" },
      callbacks,
      vi.fn(async () => {})
    );

    expect(turn.items[0]).toMatchObject({
      type: "tool_call",
      id: "call-2",
      toolName: "range.read",
      arguments: { sheetName: "Sheet1", range: "A1" },
      status: "completed",
    });
    expect(turn.items[1]).toMatchObject({
      type: "tool_result",
      toolCallId: "call-2",
      result: { values: [[1]] },
      isError: false,
    });
  });

  it("preserves falsy successful tool result data", async () => {
    const turn = createTurn();
    const execute = vi.fn(async () => ({ success: true, data: false }));

    await processToolCalls(
      [{ id: "call-falsy", name: "selection.get", arguments: "{}" }],
      new Map(),
      turn,
      new Map([["selection.get", { name: "selection.get", execute }]]),
      { permissionMode: "confirm_all" },
      createCallbacks(),
      vi.fn(async () => {})
    );

    expect(turn.items[1]).toMatchObject({
      type: "tool_result",
      toolCallId: "call-falsy",
      result: false,
      isError: false,
    });
  });

  it("logs the structured tool execution result after a tool completes", async () => {
    const turn = createTurn();
    const execute = vi.fn(async () => ({ success: true, data: { rows: 1 } }));
    const logToolExecution = vi.fn(async () => {});

    await (processToolCalls as any)(
      [{ id: "call-log", name: "range.read", arguments: "{\"sheetName\":\"Sheet1\",\"range\":\"A1\"}" }],
      new Map(),
      turn,
      new Map([["range.read", { name: "range.read", execute }]]),
      { permissionMode: "confirm_all" },
      createCallbacks(),
      vi.fn(async () => {}),
      logToolExecution
    );

    expect(logToolExecution).toHaveBeenCalledWith(expect.objectContaining({
      threadId: "thread-1",
      turnId: "turn-1",
      toolCallId: "call-log",
      toolName: "range.read",
      status: "success",
      resultSummary: summarizeForLog({ rows: 1 }),
      argumentsSummary: summarizeForLog({ sheetName: "Sheet1", range: "A1" }),
    }));
    const logCalls = logToolExecution.mock.calls as unknown as Array<[{ durationMs: number }]>;
    const loggedRecord = logCalls[0]?.[0];
    expect(loggedRecord?.durationMs).toBeGreaterThanOrEqual(0);
  });

  it("rejects hidden advanced Excel operations before approval or execution", async () => {
    const turn = createTurn();
    turn.items.push({
      type: "user_message",
      id: "user-simple",
      content: "把 A1:C20 写入数据并汇总",
      timestamp: 1,
    });
    const execute = vi.fn(async () => ({ success: true }));
    const callbacks = createCallbacks();

    await processToolCalls(
      [{
        id: "call-hidden-pivot",
        name: "office.action.apply",
        arguments: JSON.stringify({
          app: "excel",
          action: "insert",
          operation: "createPivotTable",
          filePath: "C:\\books\\a.xlsx",
          target: "range:Sheet1!A1:D20",
          params: {
            advancedIntent: "interactive-pivot",
            rowFields: ["部门"],
          },
        }),
      }],
      new Map(),
      turn,
      new Map([["office.action.apply", { name: "office.action.apply", execute }]]),
      { permissionMode: "confirm_all" },
      callbacks,
      vi.fn(async () => {}),
    );

    expect(execute).not.toHaveBeenCalled();
    expect(turn.items).toEqual(expect.arrayContaining([
      expect.objectContaining({
        type: "tool_result",
        isError: true,
        result: expect.stringContaining("工具参数校验失败"),
      }),
    ]));
  });

  it("allows the same advanced operation when the user explicitly requested it", async () => {
    const turn = createTurn();
    turn.items.push({
      type: "user_message",
      id: "user-pivot",
      content: "创建交互式数据透视表",
      timestamp: 1,
    });
    const execute = vi.fn(async () => ({ success: true, data: { created: true } }));

    await processToolCalls(
      [{
        id: "call-visible-pivot",
        name: "office.action.apply",
        arguments: JSON.stringify({
          app: "excel",
          action: "insert",
          operation: "createPivotTable",
          filePath: "C:\\books\\a.xlsx",
          target: "range:Sheet1!A1:D20",
          params: {
            advancedIntent: "interactive-pivot",
            rowFields: ["部门"],
          },
        }),
      }],
      new Map(),
      turn,
      new Map([["office.action.apply", { name: "office.action.apply", execute }]]),
      { permissionMode: "confirm_all" },
      createCallbacks(),
      vi.fn(async () => {}),
    );

    expect(execute).toHaveBeenCalledOnce();
  });

  it("stores always-allowed tools after approval requests opt in", async () => {
    const activeItem: ToolCallItem = {
      type: "tool_call",
      id: "call-3",
      toolName: "office.action.apply",
      arguments: { app: "excel", action: "style", operation: "formatChart", filePath: "C:\\books\\a.xlsx" },
      status: "pending",
      timestamp: 1000,
    };
    const toolCalls: ToolCallInfo[] = [
      {
        id: "call-3",
        name: "office.action.apply",
        arguments: "{\"app\":\"excel\",\"action\":\"style\",\"operation\":\"formatChart\",\"filePath\":\"C:\\\\books\\\\a.xlsx\"}",
      },
    ];

    await processToolCalls(
      toolCalls,
      new Map([["call-3", activeItem]]),
      createTurn(activeItem),
      new Map([["office.action.apply", { name: "office.action.apply", execute: vi.fn(async () => ({ success: true, data: "ok" })) }]]),
      {
        permissionMode: "normal",
        requestToolApproval: vi.fn(async () => ({ approved: true, alwaysAllow: true })),
      },
      createCallbacks(),
      vi.fn(async () => {})
    );

    expect(getAlwaysAllowedTools().has("office.action.apply")).toBe(true);
  });

  it("canonicalizes OCR tool aliases before execution and event storage", async () => {
    const turn = createTurn();
    const execute = vi.fn(async () => ({ success: true, data: { text: "hello" } }));

    await processToolCalls(
      [
        {
          id: "call-ocr",
          name: "ocr_parseDocument",
          arguments: "{\"filePaths\":[\"C:\\\\Users\\\\29721\\\\Pictures\\\\image.png\"]}",
        },
      ],
      new Map(),
      turn,
      new Map([["ocr.parseDocument", { name: "ocr.parseDocument", execute }]]),
      {
        permissionMode: "confirm_all",
        requestToolApproval: vi.fn(async () => ({ approved: true })),
      },
      createCallbacks(),
      vi.fn(async () => {})
    );

    expect(execute).toHaveBeenCalledWith({
      filePaths: ["C:\\Users\\29721\\Pictures\\image.png"],
    }, {
      threadId: "thread-1",
      turnId: "turn-1",
      userMessages: [],
    });
    expect(turn.items[0]).toMatchObject({
      type: "tool_call",
      toolName: "ocr.parseDocument",
      status: "completed",
    });
    expect(turn.items[1]).toMatchObject({
      type: "tool_result",
      toolName: "ocr.parseDocument",
      result: { text: "hello" },
      isError: false,
    });
  });

  it("checks interruption before starting the next tool in a batch", async () => {
    const turn = createTurn();
    let aborted = false;
    const firstExecute = vi.fn(async () => {
      aborted = true;
      return { success: true, data: "first" };
    });
    const secondExecute = vi.fn(async () => ({ success: true, data: "second" }));
    const throwIfAborted = vi.fn(() => {
      if (!aborted) return;
      const error = new Error("aborted");
      error.name = "AbortError";
      throw error;
    });

    await expect(processToolCalls(
      [
        { id: "call-first", name: "range.read", arguments: "{\"sheetName\":\"Sheet1\",\"range\":\"A1\"}" },
        { id: "call-second", name: "selection.get", arguments: "{}" },
      ],
      new Map(),
      turn,
      new Map([
        ["range.read", { name: "range.read", execute: firstExecute }],
        ["selection.get", { name: "selection.get", execute: secondExecute }],
      ]),
      { permissionMode: "confirm_all" },
      createCallbacks(),
      vi.fn(async () => {}),
      undefined,
      throwIfAborted
    )).rejects.toMatchObject({ name: "AbortError" });

    expect(firstExecute).toHaveBeenCalledOnce();
    expect(secondExecute).not.toHaveBeenCalled();
  });
});
