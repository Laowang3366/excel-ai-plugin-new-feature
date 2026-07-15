import { describe, expect, it } from "vitest";

import {
  resolveWorkflowStep,
  resolveWorkflowVariables,
  shouldRunWorkflowStep,
  type WorkflowResultReference,
} from "./workflowStepExecution";

const baseStep = {
  app: "word" as const,
  action: "edit" as const,
  operation: "formatLongDocument",
  filePath: "{{vars.files.0}}",
  params: { title: "{{vars.customer.name}}" },
};

describe("Office workflow placeholder paths", () => {
  it("resolves bounded own-property object and array paths", () => {
    const [resolved] = resolveWorkflowVariables([baseStep], {
      files: ["C:/reports/source.docx"],
      customer: { name: "示例客户" },
    });

    expect(resolved.filePath).toBe("C:/reports/source.docx");
    expect(resolved.params).toEqual({ title: "示例客户" });
  });

  it.each([
    "{{vars.constructor.name}}",
    "{{vars.customer.__proto__.name}}",
    "{{vars.customer..name}}",
    `{{vars.${Array.from({ length: 33 }, () => "item").join(".")}}}`,
  ])("rejects unsafe or ambiguous variable path %s", (filePath) => {
    expect(() =>
      resolveWorkflowVariables([{ ...baseStep, filePath }], {
        customer: { name: "示例客户" },
      }),
    ).toThrow("工作流变量没有值");
  });

  it("does not read inherited properties from step results or conditions", () => {
    const records: WorkflowResultReference[] = [
      {
        step: 1,
        id: "source",
        status: "done",
        result: {
          status: "done",
          engine: "openxml",
          app: "excel",
          action: "inspect",
          operation: "inspectFile",
          outputPath: "C:/reports/source.xlsx",
          summary: "done",
          changes: [],
          data: { ready: true },
        },
      },
    ];

    expect(
      resolveWorkflowStep(
        { ...baseStep, filePath: "{{steps.source.outputPath}}", params: {} },
        records,
      ).filePath,
    ).toBe("C:/reports/source.xlsx");
    expect(() =>
      resolveWorkflowStep(
        { ...baseStep, filePath: "{{steps.source.constructor.name}}", params: {} },
        records,
      ),
    ).toThrow("工作流占位符没有值");
    expect(
      shouldRunWorkflowStep(
        { step: "source", dataPath: "constructor.name", equals: "Object" },
        records,
      ),
    ).toBe(false);
  });
});
