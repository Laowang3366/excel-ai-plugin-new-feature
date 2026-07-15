import { describe, expect, it } from "vitest";

import { ALL_TOOL_DEFINITIONS } from "./toolDefinitions";
import { filterToolDefinitionsForTurn } from "./officeToolVisibility";

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
