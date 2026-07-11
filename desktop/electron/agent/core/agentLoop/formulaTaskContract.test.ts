import { describe, expect, it } from "vitest";
import type { Turn, TurnItem } from "../../shared/types";
import {
  getMissingRequiredReads,
  normalizeFormulaPreparation,
  parseFormulaRangeRef,
  parseFormulaTaskContract,
  rangeRefContains,
} from "./formulaTaskContract";

function turn(items: TurnItem[]): Turn {
  return { threadId: "t", turnId: "r", status: "in_progress", startedAt: 1, items };
}

function user(content: string): TurnItem {
  return { type: "user_message", id: "u", content, timestamp: 1 };
}

function read(id: string, range: string, matrix: unknown[][]): TurnItem[] {
  return [
    { type: "tool_call", id, toolName: "range.read", arguments: { sheetName: "数据表", range }, status: "completed", timestamp: 2 },
    { type: "tool_result", id: `result-${id}`, toolCallId: id, toolName: "range.read", result: matrix, isError: false, timestamp: 3 },
  ];
}

describe("formulaTaskContract", () => {
  it("parses complete, partial, and absent reference samples", () => {
    const complete = parseFormulaTaskContract(turn([user([
      "【功能模块：生成公式】",
      "任务说明：汇总",
      "数据源选区：数据表!A1:C10；辅助表!A1:B5",
      "答案参考样例：数据表!E1:F4",
      "答案参考样例类型：完整样例",
      "答案填入锚点/选区：数据表!H1",
    ].join("\n"))]));
    const none = parseFormulaTaskContract(turn([user([
      "【功能模块：生成公式】",
      "任务说明：汇总",
      "数据源选区：数据表!A1:C10",
      "答案填入锚点/选区：由 Agent 选择空白区域",
    ].join("\n"))]));

    expect(complete).toMatchObject({
      referenceMode: "complete",
      targetChosenByAgent: false,
      dataSourceRanges: [{ raw: "数据表!A1:C10" }, { raw: "辅助表!A1:B5" }],
      referenceRange: { raw: "数据表!E1:F4" },
    });
    expect(none).toMatchObject({ referenceMode: "none", targetChosenByAgent: true });
    expect(none?.referenceRange).toBeUndefined();
  });

  it("accepts a read range that contains the requested source range", () => {
    const contract = parseFormulaTaskContract(turn([user([
      "【功能模块：生成公式】",
      "数据源选区：数据表!A2:C10",
      "答案填入锚点/选区：数据表!H1",
    ].join("\n"))]))!;
    const current = turn([
      user(contract.sourceContent),
      ...read("read-wide", "A1:D20", [["部门", "金额"]]),
    ]);

    expect(getMissingRequiredReads(current, contract)).toEqual([]);
    expect(rangeRefContains(parseFormulaRangeRef("数据表!A1:D20")!, parseFormulaRangeRef("数据表!A2:C10")!)).toBe(true);
  });

  it("requires only a source read when no reference sample is configured", () => {
    const content = [
      "【功能模块：生成公式】",
      "数据源选区：数据表!A1:C10",
      "答案填入锚点/选区：数据表!H1",
    ].join("\n");
    const contract = parseFormulaTaskContract(turn([user(content)]))!;
    const current = turn([user(content), ...read("read-source", "A1:C10", [[1]])]);

    expect(contract.referenceMode).toBe("none");
    expect(getMissingRequiredReads(current, contract)).toEqual([]);
  });

  it("normalizes ready and clarification preparations", () => {
    const ready = normalizeFormulaPreparation({
      status: "ready",
      scenario: "分组聚合",
      inputShape: "记录表",
      outputShape: "汇总表",
      inputGrain: "明细",
      outputGrain: "部门",
      businessKeys: ["部门"],
      transformChain: ["分组", "聚合"],
      constraints: [],
      acceptanceChecks: [{ type: "unique_key", description: "部门唯一", params: { outputColumns: [1] } }],
    });
    const clarification = normalizeFormulaPreparation({
      status: "needs_clarification",
      clarificationQuestion: "重复数据是否计入？",
    });

    expect(ready).toMatchObject({ status: "ready", acceptanceChecks: [{ required: true }] });
    expect(clarification).toMatchObject({ status: "needs_clarification", clarificationQuestion: "重复数据是否计入？" });
  });
});
