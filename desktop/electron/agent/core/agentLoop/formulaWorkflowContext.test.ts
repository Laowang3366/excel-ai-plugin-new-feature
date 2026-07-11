import { describe, expect, it } from "vitest";
import type { Turn, TurnItem } from "../../shared/types";
import { isFormulaWorkflowTurn, parseFormulaTaskContract } from "./formulaTaskContract";
import {
  createFormulaWorkflowTurnView,
  hasPendingFormulaWorkflowHistory,
} from "./formulaWorkflowContext";

function user(id: string, content: string): TurnItem {
  return { type: "user_message", id, content, timestamp: 1 };
}

function toolResult(id: string, toolName: string, value: unknown): TurnItem[] {
  return [
    { type: "tool_call", id, toolName, arguments: {}, status: "completed", timestamp: 2 },
    { type: "tool_result", id: `result-${id}`, toolCallId: id, toolName, result: value, isError: false, timestamp: 3 },
  ];
}

const task = user("task", [
  "【功能模块：生成公式】",
  "任务说明：按部门汇总",
  "数据源选区：Sheet1!A1:B10",
  "答案填入锚点/选区：Sheet1!D1",
].join("\n"));

const clarification = {
  status: "needs_clarification",
  scenario: "分组聚合",
  inputShape: "记录表",
  outputShape: "汇总表",
  inputGrain: "明细",
  outputGrain: "部门",
  businessKeys: ["部门"],
  transformChain: [],
  constraints: [],
  acceptanceChecks: [],
  assumptions: [],
  clarificationQuestion: "重复项是否合计？",
};

function currentTurn(items: TurnItem[]): Turn {
  return { threadId: "thread", turnId: "current", status: "in_progress", startedAt: 10, items };
}

describe("formulaWorkflowContext", () => {
  it("continues a clarification workflow when the user replies without the module marker", () => {
    const previous = [task, ...toolResult("prepare-old", "formula.prepare", clarification)];
    const current = currentTurn([user("answer", "重复项需要合计")]);
    const view = createFormulaWorkflowTurnView(current, [previous, current.items]);

    expect(view).not.toBe(current);
    expect(isFormulaWorkflowTurn(view)).toBe(true);
    expect(parseFormulaTaskContract(view)?.task).toBe("按部门汇总");
    expect(hasPendingFormulaWorkflowHistory([previous, current.items])).toBe(true);
  });

  it("does not revive an older clarification after a continuation passes verification", () => {
    const previous = [task, ...toolResult("prepare-old", "formula.prepare", clarification)];
    const continuation = [
      user("answer", "重复项需要合计"),
      ...toolResult("prepare-new", "formula.prepare", { ...clarification, status: "ready", clarificationQuestion: undefined }),
      ...toolResult("verify-new", "formula.verify", {
        status: "passed",
        writeToolCallId: "write-new",
        referenceMode: "none",
        anchor: "Sheet1!D1",
        actualRange: "Sheet1!D1:E3",
        actualShape: { rows: 3, columns: 2 },
        checks: [],
        errorCells: [],
        sampleMismatches: [],
        assumptions: [],
        summary: "通过",
        nextActions: [],
      }),
    ];
    const current = currentTurn([user("next", "现在帮我写一封邮件")]);

    expect(hasPendingFormulaWorkflowHistory([previous, continuation, current.items], current.items)).toBe(false);
    expect(createFormulaWorkflowTurnView(current, [previous, continuation, current.items])).toBe(current);
  });
});
