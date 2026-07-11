import { describe, expect, it, vi } from "vitest";
import type { ToolExecutor, Turn, TurnItem } from "../../shared/types";
import {
  buildFormulaKnowledgeQuery,
  guardFormulaRangeWriteExecutors,
  hasFormulaKnowledgeSearchAttempt,
  isFormulaTurn,
  shouldAutoSearchFormulaKnowledge,
} from "./formulaKnowledgePolicy";

function createTurn(items: TurnItem[]): Turn {
  return {
    threadId: "thread-1",
    turnId: "turn-1",
    status: "in_progress",
    startedAt: 1,
    items,
  };
}

function user(content: string): TurnItem {
  return { type: "user_message", id: "u1", content, timestamp: 1 };
}

function result(toolName: string, data: unknown, isError = false): TurnItem {
  return {
    type: "tool_result",
    id: `result-${toolName}`,
    toolCallId: `call-${toolName}`,
    toolName,
    result: data,
    isError,
    timestamp: 2,
  };
}

function executor(name: string, execute = vi.fn(async () => ({ success: true }))): ToolExecutor {
  return { name, execute };
}

describe("formulaKnowledgePolicy", () => {
  it("requires automatic methodology search after formula module data is read", () => {
    const turn = createTurn([
      user("【功能模块：生成公式】\n任务说明：按部门统计剩余餐费"),
      result("range.read", { address: "B3:E12", values: [["部门", "人数"]] }),
    ]);
    const executors = new Map([
      ["knowledge.search", executor("knowledge.search")],
      ["range.write", executor("range.write")],
    ]);

    expect(isFormulaTurn(turn)).toBe(true);
    expect(shouldAutoSearchFormulaKnowledge(turn, executors)).toBe(true);
    expect(buildFormulaKnowledgeQuery(turn)).toContain("按部门统计剩余餐费");
    expect(buildFormulaKnowledgeQuery(turn)).toContain("B3:E12");
    expect(buildFormulaKnowledgeQuery(turn)).toContain("只返回与当前任务结构、约束和验收直接相关");
    expect(buildFormulaKnowledgeQuery(turn)).not.toContain("SCAN");
    expect(buildFormulaKnowledgeQuery(turn)).not.toContain("正则");
  });

  it("does not search when a formula question has not read workbook structure", () => {
    const turn = createTurn([
      user("请生成公式：SUMIFS 怎么用？"),
    ]);
    const executors = new Map([["knowledge.search", executor("knowledge.search")]]);

    expect(isFormulaTurn(turn)).toBe(true);
    expect(shouldAutoSearchFormulaKnowledge(turn, executors)).toBe(false);
  });

  it("searches methodology for non-module formula turns after structure is read", () => {
    const turn = createTurn([
      user("请读取当前表格并写入公式"),
      result("range.read", { address: "A1:D20" }),
    ]);
    const executors = new Map([["knowledge.search", executor("knowledge.search")]]);

    expect(isFormulaTurn(turn)).toBe(true);
    expect(shouldAutoSearchFormulaKnowledge(turn, executors)).toBe(true);
  });

  it("uses the same skill trigger for named formula functions", () => {
    const turn = createTurn([
      user("读取当前表格，用 SUMIFS 按部门汇总"),
      result("range.read", { address: "A1:D20" }),
    ]);
    const executors = new Map([["knowledge.search", executor("knowledge.search")]]);

    expect(isFormulaTurn(turn)).toBe(true);
    expect(shouldAutoSearchFormulaKnowledge(turn, executors)).toBe(true);
  });

  it("blocks formula writes until a knowledge search has been attempted", async () => {
    const write = vi.fn(async () => ({ success: true, data: "written" }));
    const turn = createTurn([user("【功能模块：生成公式】\n任务说明：动态汇总")]);
    const executors = new Map([
      ["knowledge.search", executor("knowledge.search")],
      ["range.write", executor("range.write", write)],
    ]);
    const guarded = guardFormulaRangeWriteExecutors(executors, turn);

    const blocked = await guarded.get("range.write")!.execute({});
    expect(blocked.success).toBe(false);
    expect(write).not.toHaveBeenCalled();

    turn.items.push(result("knowledge.search", "找到方法论", true));
    expect(hasFormulaKnowledgeSearchAttempt(turn)).toBe(true);
    const allowed = await guarded.get("range.write")!.execute({});
    expect(allowed).toEqual({ success: true, data: "written" });
    expect(write).toHaveBeenCalledOnce();
  });
});
