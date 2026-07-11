import { describe, expect, it } from "vitest";
import { addFormulaWorkflowExecutors } from "./formulaWorkflowExecutors";

describe("formula workflow executors", () => {
  it("normalizes a ready preparation and defaults checks to required", async () => {
    const executors = new Map();
    addFormulaWorkflowExecutors(executors);

    const result = await executors.get("formula.prepare").execute({
      status: "ready",
      scenario: "分组聚合",
      inputShape: "记录表",
      outputShape: "汇总表",
      inputGrain: "明细",
      outputGrain: "部门",
      businessKeys: ["部门"],
      transformChain: ["分组", "聚合"],
      constraints: [],
      acceptanceChecks: [{ type: "unique_key", description: "部门唯一", params: {} }],
    });

    expect(result).toMatchObject({
      success: true,
      data: { status: "ready", acceptanceChecks: [{ type: "unique_key", required: true }] },
    });
  });

  it("accepts a clarification question without fabricating a ready plan", async () => {
    const executors = new Map();
    addFormulaWorkflowExecutors(executors);

    const result = await executors.get("formula.prepare").execute({
      status: "needs_clarification",
      clarificationQuestion: "重复项是否计入？",
    });

    expect(result).toMatchObject({ success: true, data: { status: "needs_clarification" } });
  });

  it("rejects an incomplete ready preparation", async () => {
    const executors = new Map();
    addFormulaWorkflowExecutors(executors);

    const result = await executors.get("formula.prepare").execute({ status: "ready" });

    expect(result).toMatchObject({ success: false });
    expect(result.error).toContain("缺少结构判断字段");
  });
});
