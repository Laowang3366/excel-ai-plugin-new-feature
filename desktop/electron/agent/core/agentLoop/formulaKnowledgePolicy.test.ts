import { describe, expect, it, vi } from "vitest";
import type { ToolExecutor, Turn, TurnItem } from "../../shared/types";
import {
  buildFormulaKnowledgeQuery,
  buildFormulaSceneQuery,
  getFormulaCompletionDiagnostic,
  guardFormulaWorkflowExecutors,
  shouldLoadFormulaMethodology,
  shouldSearchFormulaScene,
} from "./formulaKnowledgePolicy";
import type { FormulaPreparation } from "./formulaTaskContract";

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

function call(id: string, toolName: string, args: Record<string, unknown>, status: "pending" | "running" | "completed" | "failed" = "completed"): TurnItem {
  return { type: "tool_call", id, toolName, arguments: args, status, timestamp: 2 };
}

function result(id: string, toolName: string, data: unknown, isError = false): TurnItem {
  return {
    type: "tool_result",
    id: `result-${id}`,
    toolCallId: id,
    toolName,
    result: data,
    isError,
    timestamp: 3,
  };
}

function executor(name: string, execute?: ToolExecutor["execute"]): ToolExecutor {
  return {
    name,
    execute: execute ?? vi.fn(async (args: Record<string, unknown>) => ({ success: true, data: args })),
  };
}

const formulaPayload = [
  "【功能模块：生成公式】",
  "任务说明：按部门汇总销售额并降序排列",
  "当前连接环境：WPS",
  "数据源选区：Sheet1!A1:C10",
  "答案参考样例：Sheet1!E1:F4",
  "答案参考样例类型：完整样例",
  "答案填入锚点/选区：Sheet1!H1",
].join("\n");

const noSamplePayload = [
  "【功能模块：生成公式】",
  "任务说明：按部门汇总销售额并降序排列",
  "数据源选区：Sheet1!A1:C10",
  "答案填入锚点/选区：Sheet1!H1",
].join("\n");

const preparation: FormulaPreparation = {
  status: "ready",
  scenario: "分组聚合",
  inputShape: "10行3列记录表",
  outputShape: "4行2列汇总表",
  inputGrain: "每行一笔销售记录",
  outputGrain: "每个部门一行",
  businessKeys: ["部门"],
  transformChain: ["按部门分组", "汇总销售额", "按金额降序"],
  constraints: ["WPS", "动态数组"],
  acceptanceChecks: [
    { type: "no_excel_error", description: "不得出现公式错误", required: true, params: {} },
    { type: "unique_key", description: "部门不得重复", required: true, params: { outputColumns: [1] } },
  ],
  assumptions: [],
};

function requiredReads(): TurnItem[] {
  return [
    call("read-source", "range.read", { sheetName: "Sheet1", range: "A1:C10" }),
    result("read-source", "range.read", [["部门", "销售额", "日期"], ["华东", 100, 1]]),
    call("read-reference", "range.read", { sheetName: "Sheet1", range: "E1:F4" }),
    result("read-reference", "range.read", [["部门", "销售额"], ["华东", 100]]),
  ];
}

function preparedItems(value: FormulaPreparation = preparation): TurnItem[] {
  return [
    call("prepare", "formula.prepare", value as unknown as Record<string, unknown>),
    result("prepare", "formula.prepare", value),
  ];
}

function knowledgeItems(scope: "formula_methodology" | "formula_scene", data: unknown, isError = false): TurnItem[] {
  const id = scope === "formula_methodology" ? "methodology" : "scene";
  return [
    call(id, "knowledge.search", { query: "结构化查询", scope }),
    result(id, "knowledge.search", data, isError),
  ];
}

describe("formulaKnowledgePolicy", () => {
  it("requires every specified source and reference range before formula.prepare", async () => {
    const turn = createTurn([
      user(formulaPayload),
      call("read-source", "range.read", { sheetName: "Sheet1", range: "A1:C10" }),
      result("read-source", "range.read", [["部门", "销售额"]]),
    ]);
    const guarded = guardFormulaWorkflowExecutors(new Map([
      ["formula.prepare", executor("formula.prepare")],
    ]), turn);

    const blocked = await guarded.get("formula.prepare")!.execute(preparation as unknown as Record<string, unknown>);

    expect(blocked.success).toBe(false);
    expect(blocked.error).toContain("REQUIRED_RANGE_NOT_READ");
    expect(blocked.error).toContain("Sheet1!E1:F4");
  });

  it("rejects preparation emitted in the same model round as the reads", async () => {
    const turn = createTurn([
      user(formulaPayload),
      call("read-source", "range.read", { sheetName: "Sheet1", range: "A1:C10" }),
      call("read-reference", "range.read", { sheetName: "Sheet1", range: "E1:F4" }),
      call("prepare", "formula.prepare", preparation as unknown as Record<string, unknown>, "running"),
      result("read-source", "range.read", [["部门", "销售额", "日期"]]),
      result("read-reference", "range.read", [["部门", "销售额"]]),
    ]);
    const guarded = guardFormulaWorkflowExecutors(new Map([
      ["formula.prepare", executor("formula.prepare")],
    ]), turn);

    const blocked = await guarded.get("formula.prepare")!.execute(preparation as unknown as Record<string, unknown>);

    expect(blocked.success).toBe(false);
    expect(blocked.error).toContain("PREPARATION_BEFORE_READ_RESULTS");
  });

  it("requires a business or structural acceptance check when no sample exists", async () => {
    const weak = {
      ...preparation,
      acceptanceChecks: [
        { type: "no_excel_error", description: "不得出现公式错误", required: true, params: {} },
      ],
    };
    const turn = createTurn([
      user(noSamplePayload),
      call("read-source", "range.read", { sheetName: "Sheet1", range: "A1:C10" }),
      result("read-source", "range.read", [["部门", "销售额"]]),
    ]);
    const guarded = guardFormulaWorkflowExecutors(new Map([
      ["formula.prepare", executor("formula.prepare")],
    ]), turn);

    const blocked = await guarded.get("formula.prepare")!.execute(weak as unknown as Record<string, unknown>);

    expect(blocked.success).toBe(false);
    expect(blocked.error).toContain("NO_SAMPLE_ACCEPTANCE_INSUFFICIENT");
  });

  it("loads mandatory methodology and optional scene knowledge only after preparation", () => {
    const executors = new Map([["knowledge.search", executor("knowledge.search")]]);
    const before = createTurn([user(formulaPayload), ...requiredReads()]);
    expect(shouldLoadFormulaMethodology(before, executors)).toBe(false);

    const ready = createTurn([user(formulaPayload), ...requiredReads(), ...preparedItems()]);
    expect(shouldLoadFormulaMethodology(ready, executors)).toBe(true);
    expect(shouldSearchFormulaScene(ready, executors)).toBe(true);
    expect(buildFormulaKnowledgeQuery(ready)).toContain("场景：分组聚合");
    expect(buildFormulaKnowledgeQuery(ready)).toContain("输出形状：4行2列汇总表");
    expect(buildFormulaSceneQuery(ready)).toContain("业务键：部门");
  });

  it("does not treat a failed methodology search as permission to write", async () => {
    const write = vi.fn(async () => ({ success: true, data: "written" }));
    const turn = createTurn([
      user(formulaPayload),
      ...requiredReads(),
      ...preparedItems(),
      ...knowledgeItems("formula_methodology", "检索失败", true),
    ]);
    const guarded = guardFormulaWorkflowExecutors(new Map([
      ["range.write", executor("range.write", write)],
      ["formula.prepare", executor("formula.prepare")],
      ["formula.verify", executor("formula.verify")],
    ]), turn);

    const blocked = await guarded.get("range.write")!.execute({ sheetName: "Sheet1", range: "H1", values: [["=SUM(A1:A2)"]] });

    expect(blocked.success).toBe(false);
    expect(blocked.error).toContain("METHODOLOGY_NOT_LOADED");
    expect(write).not.toHaveBeenCalled();
  });

  it("allows no_match scene knowledge but enforces the requested target anchor", async () => {
    const write = vi.fn(async () => ({ success: true, data: "written" }));
    const turn = createTurn([
      user(formulaPayload),
      ...requiredReads(),
      ...preparedItems(),
      ...knowledgeItems("formula_methodology", "方法论"),
      ...knowledgeItems("formula_scene", { status: "no_match", matchCount: 0 }),
    ]);
    const guarded = guardFormulaWorkflowExecutors(new Map([
      ["range.write", executor("range.write", write)],
      ["formula.prepare", executor("formula.prepare")],
      ["formula.verify", executor("formula.verify")],
    ]), turn);

    const wrongTarget = await guarded.get("range.write")!.execute({ sheetName: "Sheet1", range: "J1", values: [["=SUM(A1:A2)"]] });
    expect(wrongTarget.success).toBe(false);
    expect(wrongTarget.error).toContain("TARGET_RANGE_MISMATCH");

    const allowed = await guarded.get("range.write")!.execute({ sheetName: "Sheet1", range: "H1", values: [["=SUM(A1:A2)"]] });
    expect(allowed).toEqual({ success: true, data: "written" });
    expect(write).toHaveBeenCalledOnce();
  });

  it("blocks multiple dynamic-array anchors", async () => {
    const turn = createTurn([
      user(formulaPayload),
      ...requiredReads(),
      ...preparedItems(),
      ...knowledgeItems("formula_methodology", "方法论"),
    ]);
    const guarded = guardFormulaWorkflowExecutors(new Map([
      ["range.write", executor("range.write")],
      ["formula.prepare", executor("formula.prepare")],
      ["formula.verify", executor("formula.verify")],
    ]), turn);

    const blocked = await guarded.get("range.write")!.execute({
      sheetName: "Sheet1",
      range: "H1",
      values: [["=FILTER(A:A,A:A<>\"\")"], ["=UNIQUE(A:A)"]],
    });

    expect(blocked.success).toBe(false);
    expect(blocked.error).toContain("MULTIPLE_DYNAMIC_ARRAY_ANCHORS");
  });

  it("invalidates a preparation when newer workbook structure is read", async () => {
    const turn = createTurn([
      user(formulaPayload),
      ...requiredReads(),
      ...preparedItems(),
      ...knowledgeItems("formula_methodology", "方法论"),
      call("read-latest", "range.read", { sheetName: "Sheet1", range: "A1:C20" }),
      result("read-latest", "range.read", [["部门", "销售额", "日期"]]),
    ]);
    const guarded = guardFormulaWorkflowExecutors(new Map([
      ["range.write", executor("range.write")],
      ["formula.prepare", executor("formula.prepare")],
      ["formula.verify", executor("formula.verify")],
    ]), turn);

    const blocked = await guarded.get("range.write")!.execute({ sheetName: "Sheet1", range: "H1", values: [["=SUM(A1:A2)"]] });

    expect(blocked.success).toBe(false);
    expect(blocked.error).toContain("FORMULA_PREPARATION_STALE");
  });

  it("does not reuse a passed verification after a newer write", () => {
    const turn = createTurn([
      user(noSamplePayload),
      call("read-source", "range.read", { sheetName: "Sheet1", range: "A1:C10" }),
      result("read-source", "range.read", [["部门", "销售额"]]),
      ...preparedItems(),
      ...knowledgeItems("formula_methodology", "方法论"),
      call("write-old", "range.write", { sheetName: "Sheet1", range: "H1", values: [["=SUM(A:A)"]] }),
      result("write-old", "range.write", "写入成功"),
      call("verify-old", "formula.verify", {}),
      result("verify-old", "formula.verify", {
        status: "passed",
        writeToolCallId: "write-old",
        referenceMode: "none",
        anchor: "Sheet1!H1",
        actualRange: "Sheet1!H1",
        actualShape: { rows: 1, columns: 1 },
        checks: [],
        errorCells: [],
        sampleMismatches: [],
        assumptions: [],
        summary: "通过",
        nextActions: [],
      }),
      call("write-new", "range.write", { sheetName: "Sheet1", range: "H1", values: [["=SUM(B:B)"]] }),
      result("write-new", "range.write", "写入成功"),
    ]);

    expect(getFormulaCompletionDiagnostic(turn)).toMatchObject({ code: "FORMULA_NOT_VERIFIED" });
  });

  it("allows a clarification response to finish without writing", () => {
    const clarification: FormulaPreparation = {
      ...preparation,
      status: "needs_clarification",
      clarificationQuestion: "重复部门应合计还是只取第一条？",
    };
    const turn = createTurn([user(formulaPayload), ...requiredReads(), ...preparedItems(clarification)]);

    expect(getFormulaCompletionDiagnostic(turn)).toBeNull();
  });

  it("keeps ordinary formula explanations outside the write workflow", () => {
    const turn = createTurn([user("SUMIFS 怎么用？")]);

    expect(getFormulaCompletionDiagnostic(turn)).toBeNull();
  });
});
