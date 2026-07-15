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
  it("removes advanced Excel operations from simple edit turns", () => {
    const definitions = filterToolDefinitionsForTurn(
      ALL_TOOL_DEFINITIONS,
      { content: "把 Sheet1!A1:C20 写入数据并设置格式" },
    );

    for (const operation of ADVANCED_OPERATIONS) {
      expect(getOperations(definitions, "office.action.apply")).not.toContain(operation);
      expect(getWorkflowOperations(definitions)).not.toContain(operation);
    }
    expect(getOperations(definitions, "office.action.inspect"))
      .not.toContain("inspectPowerQueries");
  });

  it("exposes only Power Query operations for explicit refreshable ETL", () => {
    const definitions = filterToolDefinitionsForTurn(
      ALL_TOOL_DEFINITIONS,
      { content: "用 Power Query 合并多个外部 CSV 并支持刷新" },
    );

    expect(getOperations(definitions, "office.action.apply"))
      .toEqual(expect.arrayContaining(["createPowerQuery", "managePowerQuery"]));
    expect(getOperations(definitions, "office.action.apply"))
      .not.toEqual(expect.arrayContaining(["createPivotTable", "addSlicer"]));
    expect(getOperations(definitions, "office.action.inspect"))
      .toContain("inspectPowerQueries");
  });

  it("exposes only pivot operations for explicit interactive pivot tasks", () => {
    const definitions = filterToolDefinitionsForTurn(
      ALL_TOOL_DEFINITIONS,
      { content: "创建数据透视表并添加切片器" },
    );

    expect(getOperations(definitions, "office.action.apply"))
      .toEqual(expect.arrayContaining(["createPivotTable", "refreshPivotTables", "addSlicer"]));
    expect(getOperations(definitions, "office.action.apply"))
      .not.toEqual(expect.arrayContaining(["createPowerQuery", "managePowerQuery"]));
    expect(getWorkflowOperations(definitions))
      .toEqual(expect.arrayContaining(["createPivotTable", "addSlicer"]));
  });

  it("uses strict operation-specific params for advanced Excel operations", () => {
    const definitions = filterToolDefinitionsForTurn(
      ALL_TOOL_DEFINITIONS,
      { content: "用 Power Query 合并多个外部 CSV 并支持刷新" },
    );
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
        mFormula: "let Source = Csv.Document(File.Contents(\"C:/sales.csv\")) in Source",
        loadMode: "worksheet",
        destination: "QueryOutput!A1",
      },
    };

    expect(parseAndValidateToolArguments(JSON.stringify(valid), apply?.parameters).error)
      .toBeUndefined();
    expect(parseAndValidateToolArguments(
      JSON.stringify({ ...valid, params: { ...valid.params, shellCommand: "whoami" } }),
      apply?.parameters,
    ).error).toContain("shellCommand");
    expect(parseAndValidateToolArguments(
      JSON.stringify({ ...valid, params: { ...valid.params, loadMode: "arbitrary" } }),
      apply?.parameters,
    ).error).toContain("loadMode");
  });

  it("keeps ordinary params compatible and applies strict schemas to workflow steps", () => {
    const ordinaryDefinitions = filterToolDefinitionsForTurn(
      ALL_TOOL_DEFINITIONS,
      { content: "给已有数据插入柱状图" },
    );
    const apply = ordinaryDefinitions.find((tool) => tool.name === "office.action.apply");
    expect(parseAndValidateToolArguments(JSON.stringify({
      app: "excel",
      action: "insert",
      operation: "insertChart",
      filePath: "C:/book.xlsx",
      target: "range:Sheet1!A1:B10",
      params: { chartType: "column" },
    }), apply?.parameters).error).toBeUndefined();
    expect(getDiscriminatedOperations(ordinaryDefinitions, "office.action.apply"))
      .not.toEqual(expect.arrayContaining(ADVANCED_OPERATIONS));

    const pivotDefinitions = filterToolDefinitionsForTurn(
      ALL_TOOL_DEFINITIONS,
      { content: "创建交互式数据透视表" },
    );
    const workflow = pivotDefinitions.find((tool) => tool.name === "office.workflow.run");
    expect(parseAndValidateToolArguments(JSON.stringify({ steps: [{
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
    }] }), workflow?.parameters).error).toContain("unexpectedField");
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
