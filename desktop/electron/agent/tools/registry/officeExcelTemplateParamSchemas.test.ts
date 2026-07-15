import { describe, expect, it } from "vitest";

import { ALL_TOOL_DEFINITIONS } from "./toolDefinitions";
import { parseAndValidateToolArguments } from "./toolSchema";

function parameters(name: string) {
  const definition = ALL_TOOL_DEFINITIONS.find((tool) => tool.name === name);
  if (!definition) throw new Error(`missing tool definition: ${name}`);
  return definition.parameters;
}

describe("Excel workbook template parameter schemas", () => {
  it("accepts only Worker-supported applyWorkbookTemplate parameters", () => {
    const step = {
      app: "excel",
      action: "style",
      operation: "applyWorkbookTemplate",
      filePath: "C:/book.xlsx",
      params: {
        host: "wps",
        preset: "financial",
        sheetNames: ["Summary", "Details"],
        allSheets: false,
        fontName: "Microsoft YaHei",
        fontSize: 10.5,
        autoFit: true,
        showGridlines: false,
        freezeRows: 2,
      },
    };

    expect(
      parseAndValidateToolArguments(JSON.stringify(step), parameters("office.action.apply")).error,
    ).toBeUndefined();
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({ steps: [step] }),
        parameters("office.workflow.run"),
      ).error,
    ).toBeUndefined();
  });

  it("rejects unsupported captured-template and nested rule parameters", () => {
    const base = {
      app: "excel",
      action: "style",
      operation: "applyWorkbookTemplate",
      filePath: "C:/book.xlsx",
    };
    const apply = parameters("office.action.apply");

    expect(
      parseAndValidateToolArguments(
        JSON.stringify({ ...base, params: { preset: "corporate" } }),
        apply,
      ).error,
    ).toContain("preset");
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({ ...base, params: { template: { version: 1, sheets: [] } } }),
        apply,
      ).error,
    ).toContain("template");
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({ ...base, params: { preset: "minimal", columnRules: [] } }),
        apply,
      ).error,
    ).toContain("columnRules");
  });

  it("uses strict empty business params for capture and inspect", () => {
    for (const operation of ["captureWorkbookTemplate", "inspectWorkbookFormatting"]) {
      const valid = {
        app: "excel",
        operation,
        filePath: "C:/book.xlsx",
        params: { host: "excel" },
      };
      const inspect = parameters("office.action.inspect");
      const validate = parameters("office.action.validate");

      expect(parseAndValidateToolArguments(JSON.stringify(valid), inspect).error).toBeUndefined();
      expect(parseAndValidateToolArguments(JSON.stringify(valid), validate).error).toBeUndefined();
      expect(
        parseAndValidateToolArguments(
          JSON.stringify({ ...valid, params: { host: "excel", headerRows: 2 } }),
          inspect,
        ).error,
      ).toContain("headerRows");
    }
  });
});
