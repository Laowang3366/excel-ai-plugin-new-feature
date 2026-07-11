import { describe, expect, it } from "vitest";
import type { Turn, TurnItem } from "../../shared/types";
import type { FormulaPreparation } from "./formulaTaskContract";
import {
  buildPendingFormulaValidationRead,
  shouldRunFormulaVerification,
  verifyLatestFormulaWrite,
} from "./formulaVerification";

function turn(items: TurnItem[]): Turn {
  return { threadId: "t", turnId: "r", status: "in_progress", startedAt: 1, items };
}

function user(content: string): TurnItem {
  return { type: "user_message", id: "u", content, timestamp: 1 };
}

function call(id: string, toolName: string, args: Record<string, unknown>): TurnItem {
  return { type: "tool_call", id, toolName, arguments: args, status: "completed", timestamp: 2 };
}

function result(id: string, toolName: string, data: unknown, isError = false): TurnItem {
  return { type: "tool_result", id: `result-${id}`, toolCallId: id, toolName, result: data, isError, timestamp: 3 };
}

const basePreparation: FormulaPreparation = {
  status: "ready",
  scenario: "分组聚合",
  inputShape: "记录表",
  outputShape: "两列汇总表",
  inputGrain: "销售明细",
  outputGrain: "每个部门一行",
  businessKeys: ["部门"],
  transformChain: ["分组", "聚合", "排序"],
  constraints: [],
  acceptanceChecks: [{ type: "no_excel_error", description: "无错误值", required: true, params: {} }],
  assumptions: [],
};

function commonItems(options: {
  referenceMode?: "none" | "partial" | "complete";
  reference?: unknown[][];
  output?: unknown[][];
  preparation?: FormulaPreparation;
} = {}): TurnItem[] {
  const referenceMode = options.referenceMode ?? "complete";
  const content = [
    "【功能模块：生成公式】",
    "任务说明：按部门汇总销售额并降序排列",
    "数据源选区：Sheet1!A1:B5",
    ...(referenceMode === "none" ? [] : [
      "答案参考样例：Sheet1!D1:E3",
      `答案参考样例类型：${referenceMode === "complete" ? "完整样例" : "部分样例"}`,
    ]),
    "答案填入锚点/选区：Sheet1!G1",
  ].join("\n");
  const items: TurnItem[] = [
    user(content),
    call("source", "range.read", { sheetName: "Sheet1", range: "A1:B5" }),
    result("source", "range.read", [["部门", "销售额"], ["甲", 70], ["乙", 30], ["甲", 30], ["", ""]]),
  ];
  if (referenceMode !== "none") {
    items.push(
      call("reference", "range.read", { sheetName: "Sheet1", range: "D1:E3" }),
      result("reference", "range.read", options.reference ?? [["部门", "销售额"], ["甲", 100], ["乙", 30]]),
    );
  }
  items.push(
    call("prepare", "formula.prepare", {}),
    result("prepare", "formula.prepare", options.preparation ?? basePreparation),
    call("write", "range.write", { sheetName: "Sheet1", range: "G1", values: [["=LET(...)" ]] }),
    result("write", "range.write", "写入成功"),
  );
  if (options.output) {
    items.push(
      call("validate-read", "range.read", { sheetName: "Sheet1", range: "G1", expand: "spill" }),
      result("validate-read", "range.read", { address: "G1:H3", expanded: true, values: options.output }),
    );
  }
  return items;
}

describe("formulaVerification", () => {
  it("requests an automatic spill read after a successful formula write", () => {
    const current = turn(commonItems());

    expect(buildPendingFormulaValidationRead(current)).toEqual({ sheetName: "Sheet1", range: "G1", expand: "spill" });
    expect(shouldRunFormulaVerification(current)).toBe(false);
  });

  it("passes a complete sample only when shape and values match", () => {
    const current = turn(commonItems({ output: [["部门", "销售额"], ["甲", 100], ["乙", 30]] }));
    const report = verifyLatestFormulaWrite(current);

    expect(typeof report).not.toBe("string");
    expect(report).toMatchObject({ status: "passed", referenceMode: "complete", actualShape: { rows: 3, columns: 2 } });
    expect(shouldRunFormulaVerification(current)).toBe(true);
  });

  it("reports exact sample mismatches with cell evidence", () => {
    const report = verifyLatestFormulaWrite(turn(commonItems({
      output: [["部门", "销售额"], ["甲", 90], ["乙", 30]],
    })));

    expect(report).toMatchObject({
      status: "failed",
      sampleMismatches: [{ row: 2, column: 2, expected: 100, actual: 90 }],
    });
  });

  it("does not treat nonnumeric text as equal to numeric zero", () => {
    const report = verifyLatestFormulaWrite(turn(commonItems({
      reference: [["部门", "销售额"], ["甲", "abc"], ["乙", 30]],
      output: [["部门", "销售额"], ["甲", 0], ["乙", 30]],
    })));

    expect(report).toMatchObject({
      status: "failed",
      sampleMismatches: [{ row: 2, column: 2, expected: "abc", actual: 0 }],
    });
  });

  it("allows a partial sample when the known top-left values match a larger output", () => {
    const report = verifyLatestFormulaWrite(turn(commonItems({
      referenceMode: "partial",
      reference: [["部门", "销售额"], ["甲", 100]],
      output: [["部门", "销售额"], ["甲", 100], ["乙", 30], ["丙", 10]],
    })));

    expect(report).toMatchObject({ status: "passed", referenceMode: "partial", sampleMismatches: [] });
  });

  it("validates no-sample output through generic invariants", () => {
    const preparation: FormulaPreparation = {
      ...basePreparation,
      acceptanceChecks: [
        { type: "no_excel_error", description: "无错误值", required: true, params: {} },
        { type: "shape", description: "输出两列", required: true, params: { expectedColumns: 2, minRows: 2 } },
        { type: "unique_key", description: "部门唯一", required: true, params: { outputColumns: [1], headerRows: 1 } },
        { type: "row_count", description: "部门数一致", required: true, params: { mode: "unique_source", sourceRange: "Sheet1!A1:B5", sourceColumns: [1], sourceHeaderRows: 1, outputHeaderRows: 1 } },
        { type: "aggregate_reconciliation", description: "总额守恒", required: true, params: { sourceRange: "Sheet1!A1:B5", sourceColumn: 2, outputColumn: 2 } },
        { type: "sort_order", description: "金额降序", required: true, params: { outputColumn: 2, direction: "desc" } },
      ],
    };
    const report = verifyLatestFormulaWrite(turn(commonItems({
      referenceMode: "none",
      preparation,
      output: [["部门", "销售额"], ["甲", 100], ["乙", 30]],
    })));

    expect(report).toMatchObject({ status: "passed", referenceMode: "none" });
    expect(typeof report === "string" ? [] : report.checks.every((check) => check.status === "passed")).toBe(true);
  });

  it("fails no-sample validation when totals do not reconcile or errors appear", () => {
    const preparation: FormulaPreparation = {
      ...basePreparation,
      acceptanceChecks: [
        { type: "aggregate_reconciliation", description: "总额守恒", required: true, params: { sourceRange: "Sheet1!A1:B5", sourceColumn: 2, outputColumn: 2 } },
      ],
    };
    const report = verifyLatestFormulaWrite(turn(commonItems({
      referenceMode: "none",
      preparation,
      output: [["部门", "销售额"], ["甲", 90], ["乙", "#VALUE!"]],
    })));

    expect(report).toMatchObject({ status: "failed", errorCells: [{ row: 3, column: 2, value: "#VALUE!" }] });
    expect(typeof report === "string" ? [] : report.checks.filter((check) => check.status === "failed").map((check) => check.type)).toEqual(
      expect.arrayContaining(["no_excel_error", "aggregate_reconciliation"]),
    );
  });

  it("executes lookup, pattern, boundary, and representative spot checks", () => {
    const preparation: FormulaPreparation = {
      ...basePreparation,
      scenario: "查找映射",
      acceptanceChecks: [
        { type: "lookup_consistency", description: "键值映射一致", required: true, params: { sourceRange: "Sheet1!A1:B5", sourceKeyColumn: 1, sourceValueColumn: 2, outputKeyColumn: 1, outputValueColumn: 2 } },
        { type: "pattern_match", description: "部门格式有效", required: true, params: { outputColumn: 1, pattern: "^(甲|乙)$" } },
        { type: "boundary", description: "输出不得为空", required: true, params: { headerRows: 1, minRows: 2, allowBlank: false } },
        { type: "spot_check", description: "抽样映射一致", required: true, params: { sourceRange: "Sheet1!A1:B5", sourceKeyColumn: 1, sourceValueColumn: 2, outputKeyColumn: 1, outputValueColumn: 2, sampleSize: 2 } },
      ],
    };
    const report = verifyLatestFormulaWrite(turn(commonItems({
      referenceMode: "none",
      preparation,
      output: [["部门", "销售额"], ["甲", 70], ["乙", 30]],
    })));

    expect(report).toMatchObject({ status: "passed" });
    expect(typeof report === "string" ? [] : report.checks.map((check) => check.type)).toEqual(
      expect.arrayContaining(["lookup_consistency", "pattern_match", "boundary", "spot_check"]),
    );
  });

  it("fails lookup validation when output contains a key absent from the source", () => {
    const preparation: FormulaPreparation = {
      ...basePreparation,
      acceptanceChecks: [
        { type: "lookup_consistency", description: "键值映射一致", required: true, params: { sourceRange: "Sheet1!A1:B5", sourceKeyColumn: 1, sourceValueColumn: 2, outputKeyColumn: 1, outputValueColumn: 2 } },
      ],
    };
    const report = verifyLatestFormulaWrite(turn(commonItems({
      referenceMode: "none",
      preparation,
      output: [["部门", "销售额"], ["丙", 20]],
    })));

    expect(report).toMatchObject({ status: "failed" });
    expect(JSON.stringify(report)).toContain("数据源中未找到");
  });
});
