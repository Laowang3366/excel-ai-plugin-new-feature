import { describe, expect, it } from "vitest";

import { ALL_TOOL_DEFINITIONS } from "./toolDefinitions";
import { parseAndValidateToolArguments } from "./toolSchema";

function parameters(name: string) {
  const definition = ALL_TOOL_DEFINITIONS.find((tool) => tool.name === name);
  if (!definition) throw new Error(`missing tool definition: ${name}`);
  return definition.parameters;
}

describe("Excel workbook object parameter schemas", () => {
  it("accepts the supported object type branches", () => {
    const apply = parameters("office.action.apply");
    const cases = [
      { objectType: "worksheet", command: "rename", sheetName: "Sheet1", newName: "Summary" },
      { objectType: "name", command: "create", name: "Revenue", refersTo: "=Summary!$B$2" },
      { objectType: "table", command: "update", name: "Sales", showTotals: true },
      { objectType: "chart", command: "update", index: 1, width: 640, height: 360 },
      { objectType: "shape", command: "delete", name: "Rectangle 1" },
      { objectType: "connection", command: "refresh", name: "SalesConnection" },
    ];

    for (const params of cases) {
      const args = {
        app: "excel",
        action: "edit",
        operation: "manageWorkbookObject",
        filePath: "C:/book.xlsx",
        target: "range:Sheet1!A1:D20",
        params,
      };
      expect(parseAndValidateToolArguments(JSON.stringify(args), apply).error).toBeUndefined();
    }
  });

  it("rejects unsupported object types, commands, and cross-branch fields", () => {
    const base = {
      app: "excel",
      action: "edit",
      operation: "manageWorkbookObject",
      filePath: "C:/book.xlsx",
    };
    const apply = parameters("office.action.apply");

    expect(
      parseAndValidateToolArguments(
        JSON.stringify({ ...base, params: { objectType: "pivotTable", command: "refresh" } }),
        apply,
      ).error,
    ).toContain("params");
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({
          ...base,
          params: { objectType: "connection", command: "update", name: "Sales" },
        }),
        apply,
      ).error,
    ).toContain("params");
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({
          ...base,
          params: { objectType: "chart", command: "update", index: 1, refersTo: "=A1" },
        }),
        apply,
      ).error,
    ).toContain("params");
  });

  it("constrains the legacy worksheet shape entry", () => {
    const args = {
      app: "excel",
      action: "edit",
      operation: "manageWorksheetObjects",
      filePath: "C:/book.xlsx",
      target: "range:Sheet1!A1",
      params: { command: "update", index: 1, left: 20, top: 30, width: 200, height: 80 },
    };
    const apply = parameters("office.action.apply");

    expect(parseAndValidateToolArguments(JSON.stringify(args), apply).error).toBeUndefined();
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({ ...args, params: { command: "add", name: "NewShape" } }),
        apply,
      ).error,
    ).toContain("command");
  });

  it("rejects the unimplemented inspect types filter", () => {
    const valid = {
      app: "excel",
      operation: "inspectWorkbookObjects",
      filePath: "C:/book.xlsx",
      params: { host: "excel" },
    };

    for (const name of ["office.action.inspect", "office.action.validate"]) {
      const schema = parameters(name);
      expect(parseAndValidateToolArguments(JSON.stringify(valid), schema).error).toBeUndefined();
      expect(
        parseAndValidateToolArguments(
          JSON.stringify({ ...valid, params: { host: "excel", types: ["chart"] } }),
          schema,
        ).error,
      ).toContain("types");
    }
  });
});
