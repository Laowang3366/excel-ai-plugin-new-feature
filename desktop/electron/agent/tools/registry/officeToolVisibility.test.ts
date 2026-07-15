import { describe, expect, it } from "vitest";

import { ALL_TOOL_DEFINITIONS } from "./toolDefinitions";
import { filterToolDefinitionsForTurn } from "./officeToolVisibility";
import { parseAndValidateToolArguments } from "./toolSchema";

const ADVANCED_OPERATIONS = [
  "createPowerQuery",
  "managePowerQuery",
  "inspectPowerQueries",
  "createPivotTable",
  "refreshPivotTables",
  "addSlicer",
];

describe("filterToolDefinitionsForTurn", () => {
  it("keeps every Office operation on an independent strict params branch", () => {
    for (const toolName of [
      "office.action.inspect",
      "office.action.apply",
      "office.action.validate",
      "office.workflow.run",
    ]) {
      const definition = ALL_TOOL_DEFINITIONS.find((tool) => tool.name === toolName);
      const variants = officeOperationVariants(definition?.parameters, toolName);
      expect(variants.length).toBeGreaterThan(0);
      for (const variant of variants) {
        const properties = variant.properties as Record<string, any> | undefined;
        expect(
          properties?.operation?.const,
          `${toolName} contains a generic operation branch`,
        ).toEqual(expect.any(String));
        expectStrictParamLeaves(properties?.params, `${toolName}.${properties?.operation?.const}`);
      }
    }
  });

  it("removes advanced Excel operations from simple edit turns", () => {
    const definitions = filterToolDefinitionsForTurn(ALL_TOOL_DEFINITIONS, {
      content: "把 Sheet1!A1:C20 写入数据并设置格式",
    });

    for (const operation of ADVANCED_OPERATIONS) {
      expect(getOperations(definitions, "office.action.apply")).not.toContain(operation);
      expect(getWorkflowOperations(definitions)).not.toContain(operation);
    }
    expect(getOperations(definitions, "office.action.inspect")).not.toContain(
      "inspectPowerQueries",
    );
  });

  it("exposes only Power Query operations for explicit refreshable ETL", () => {
    const definitions = filterToolDefinitionsForTurn(ALL_TOOL_DEFINITIONS, {
      content: "用 Power Query 合并多个外部 CSV 并支持刷新",
    });

    expect(getOperations(definitions, "office.action.apply")).toEqual(
      expect.arrayContaining(["createPowerQuery", "managePowerQuery"]),
    );
    expect(getOperations(definitions, "office.action.apply")).not.toEqual(
      expect.arrayContaining(["createPivotTable", "addSlicer"]),
    );
    expect(getOperations(definitions, "office.action.inspect")).toContain("inspectPowerQueries");
  });

  it("exposes only pivot operations for explicit interactive pivot tasks", () => {
    const definitions = filterToolDefinitionsForTurn(ALL_TOOL_DEFINITIONS, {
      content: "创建数据透视表并添加切片器",
    });

    expect(getOperations(definitions, "office.action.apply")).toEqual(
      expect.arrayContaining(["createPivotTable", "refreshPivotTables", "addSlicer"]),
    );
    expect(getOperations(definitions, "office.action.apply")).not.toEqual(
      expect.arrayContaining(["createPowerQuery", "managePowerQuery"]),
    );
    expect(getWorkflowOperations(definitions)).toEqual(
      expect.arrayContaining(["createPivotTable", "addSlicer"]),
    );
  });

  it("uses strict operation-specific params for advanced Excel operations", () => {
    const definitions = filterToolDefinitionsForTurn(ALL_TOOL_DEFINITIONS, {
      content: "用 Power Query 合并多个外部 CSV 并支持刷新",
    });
    const apply = definitions.find((tool) => tool.name === "office.action.apply");
    const valid = {
      app: "excel",
      action: "edit",
      operation: "createPowerQuery",
      filePath: "C:/book.xlsx",
      params: {
        advancedIntent: "refreshable-etl",
        sourceKind: "external",
        name: "SalesImport",
        mFormula: 'let Source = Csv.Document(File.Contents("C:/sales.csv")) in Source',
        loadMode: "worksheet",
        destination: "QueryOutput!A1",
      },
    };

    expect(
      parseAndValidateToolArguments(JSON.stringify(valid), apply?.parameters).error,
    ).toBeUndefined();
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({ ...valid, params: { ...valid.params, shellCommand: "whoami" } }),
        apply?.parameters,
      ).error,
    ).toContain("shellCommand");
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({ ...valid, params: { ...valid.params, loadMode: "arbitrary" } }),
        apply?.parameters,
      ).error,
    ).toContain("loadMode");
  });

  it("uses app-specific strict params for modeled ordinary operations", () => {
    const ordinaryDefinitions = filterToolDefinitionsForTurn(ALL_TOOL_DEFINITIONS, {
      content: "给已有数据插入柱状图",
    });
    const apply = ordinaryDefinitions.find((tool) => tool.name === "office.action.apply");
    const excelChart = {
      app: "excel",
      action: "insert",
      operation: "insertChart",
      filePath: "C:/book.xlsx",
      target: "range:Sheet1!A1:B10",
      params: { chartType: "column", host: "wps" },
    };
    expect(
      parseAndValidateToolArguments(JSON.stringify(excelChart), apply?.parameters).error,
    ).toBeUndefined();
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({
          ...excelChart,
          params: { chartType: "column", shellCommand: "whoami" },
        }),
        apply?.parameters,
      ).error,
    ).toContain("shellCommand");
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({
          ...excelChart,
          params: { chartType: "radar" },
        }),
        apply?.parameters,
      ).error,
    ).toContain("chartType");
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({
          ...excelChart,
          params: { chartType: "column", host: "powerpoint" },
        }),
        apply?.parameters,
      ).error,
    ).toContain("host");

    expect(
      parseAndValidateToolArguments(
        JSON.stringify({
          app: "presentation",
          action: "insert",
          operation: "insertChart",
          filePath: "C:/deck.pptx",
          target: "slide:2",
          params: {
            chartType: "pie",
            host: "powerpoint",
            left: 80,
            top: 120,
            width: 520,
            height: 300,
          },
        }),
        apply?.parameters,
      ).error,
    ).toBeUndefined();
    expect(getDiscriminatedOperations(ordinaryDefinitions, "office.action.apply")).not.toEqual(
      expect.arrayContaining(ADVANCED_OPERATIONS),
    );
  });

  it("applies operation-specific params and bounded variables to workflow steps", () => {
    const pivotDefinitions = filterToolDefinitionsForTurn(ALL_TOOL_DEFINITIONS, {
      content: "创建交互式数据透视表",
    });
    const workflow = pivotDefinitions.find((tool) => tool.name === "office.workflow.run");
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({
          steps: [
            {
              app: "excel",
              action: "insert",
              operation: "createPivotTable",
              filePath: "C:/book.xlsx",
              target: "range:Sheet1!A1:D20",
              params: {
                advancedIntent: "interactive-pivot",
                rowFields: ["Department"],
                unexpectedField: true,
              },
            },
          ],
        }),
        workflow?.parameters,
      ).error,
    ).toContain("unexpectedField");

    const ordinaryDefinitions = filterToolDefinitionsForTurn(ALL_TOOL_DEFINITIONS, {
      content: "把 Word 文档标题改为一级标题并更新目录",
    });
    const ordinaryWorkflow = ordinaryDefinitions.find(
      (tool) => tool.name === "office.workflow.run",
    );
    const workflowArgs = {
      variables: { heading_prefix: "第", customer: { name: "示例客户" } },
      steps: [
        {
          app: "word",
          action: "style",
          operation: "applyHeadingStyles",
          filePath: "C:/report.docx",
          params: { startsWith: "{{vars.heading_prefix}}", level: 1 },
        },
      ],
    };
    expect(
      parseAndValidateToolArguments(JSON.stringify(workflowArgs), ordinaryWorkflow?.parameters)
        .error,
    ).toBeUndefined();
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({
          ...workflowArgs,
          steps: [{ ...workflowArgs.steps[0], params: { level: 1, unknown: true } }],
        }),
        ordinaryWorkflow?.parameters,
      ).error,
    ).toContain("unknown");
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({
          ...workflowArgs,
          variables: { "customer.name": "不可使用点号作为顶层键" },
        }),
        ordinaryWorkflow?.parameters,
      ).error,
    ).toContain("格式不符合要求");
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({
          ...workflowArgs,
          variables: Object.fromEntries(
            Array.from({ length: 129 }, (_, index) => [`key_${index}`, index]),
          ),
        }),
        ordinaryWorkflow?.parameters,
      ).error,
    ).toContain("最多允许 128 个字段");
    for (const invalidStep of [
      { ...workflowArgs.steps[0], id: "bad.id" },
      { ...workflowArgs.steps[0], parallelGroup: "" },
      { ...workflowArgs.steps[0], when: { step: 0 } },
      { ...workflowArgs.steps[0], when: { step: "source", dataPath: "constructor.name" } },
    ]) {
      expect(
        parseAndValidateToolArguments(
          JSON.stringify({ ...workflowArgs, steps: [invalidStep] }),
          ordinaryWorkflow?.parameters,
        ).error,
      ).toBeDefined();
    }
  });
});

function getOperations(definitions: typeof ALL_TOOL_DEFINITIONS, toolName: string): string[] {
  const definition = definitions.find((tool) => tool.name === toolName);
  const properties = definition?.parameters.properties as Record<string, any> | undefined;
  return properties?.operation?.enum ?? [];
}

function getWorkflowOperations(definitions: typeof ALL_TOOL_DEFINITIONS): string[] {
  const definition = definitions.find((tool) => tool.name === "office.workflow.run");
  const properties = definition?.parameters.properties as Record<string, any> | undefined;
  return properties?.steps?.items?.properties?.operation?.enum ?? [];
}

function getDiscriminatedOperations(
  definitions: typeof ALL_TOOL_DEFINITIONS,
  toolName: string,
): string[] {
  const definition = definitions.find((tool) => tool.name === toolName);
  const variants = Array.isArray(definition?.parameters.oneOf) ? definition.parameters.oneOf : [];
  return variants.flatMap((variant) => {
    const operation = (variant as Record<string, any>)?.properties?.operation;
    if (typeof operation?.const === "string") return [operation.const];
    return Array.isArray(operation?.enum) ? operation.enum : [];
  });
}

function officeOperationVariants(
  parameters: Record<string, any> | undefined,
  toolName: string,
): Array<Record<string, any>> {
  const schema =
    toolName === "office.workflow.run" ? parameters?.properties?.steps?.items : parameters;
  return Array.isArray(schema?.oneOf) ? schema.oneOf : [];
}

function expectStrictParamLeaves(schema: Record<string, any> | undefined, label: string): void {
  expect(schema, `${label} is missing params schema`).toBeDefined();
  const branches = ["oneOf", "allOf"].flatMap((keyword) =>
    Array.isArray(schema?.[keyword]) ? schema[keyword] : [],
  );
  if (branches.length > 0) {
    for (const branch of branches) expectStrictParamLeaves(branch, label);
    return;
  }
  expect(schema?.type, `${label} params must end in an object schema`).toBe("object");
  expect(schema?.additionalProperties, `${label} params must reject unknown fields`).toBe(false);
}
