import { describe, expect, it } from "vitest";

import { ALL_TOOL_DEFINITIONS } from "./toolDefinitions";
import { parseAndValidateToolArguments } from "./toolSchema";

function parameters(name: string) {
  const definition = ALL_TOOL_DEFINITIONS.find((tool) => tool.name === name);
  if (!definition) throw new Error(`missing tool definition: ${name}`);
  return definition.parameters;
}

describe("Excel print operation parameter schemas", () => {
  it("accepts Worker-supported configurePrint parameters in apply and workflow", () => {
    const step = {
      app: "excel",
      action: "style",
      operation: "configurePrint",
      filePath: "C:/book.xlsx",
      params: {
        host: "wps",
        sheetNames: ["Summary", "Details"],
        orientation: "landscape",
        paperSize: "a4",
        printArea: "$A$1:$H$50",
        repeatRows: "$1:$2",
        margins: { top: 1.5, bottom: 1.5, left: 1, right: 1 },
        marginUnit: "centimeters",
        fitToOnePageWide: true,
        fitToOnePageTall: false,
        centerHorizontally: true,
        headers: { left: "Confidential", center: "Quarterly report" },
        footers: { center: "Page &P of &N" },
        clearPageBreaks: true,
        horizontalPageBreaks: ["A26"],
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

  it("rejects unsupported and unknown configurePrint parameters", () => {
    const base = {
      app: "excel",
      action: "style",
      operation: "configurePrint",
      filePath: "C:/book.xlsx",
    };
    const schema = parameters("office.action.apply");

    expect(
      parseAndValidateToolArguments(
        JSON.stringify({ ...base, params: { paperSize: "tabloid" } }),
        schema,
      ).error,
    ).toContain("paperSize");
    expect(
      parseAndValidateToolArguments(JSON.stringify({ ...base, params: { scale: 401 } }), schema)
        .error,
    ).toContain("scale");
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({ ...base, params: { margins: { top: 1, gutter: 2 } } }),
        schema,
      ).error,
    ).toContain("gutter");
  });

  it("uses strict inspectPrintSettings params for inspect and validate", () => {
    const valid = {
      app: "excel",
      operation: "inspectPrintSettings",
      filePath: "C:/book.xlsx",
      params: { host: "excel", sheetNames: ["Summary"] },
    };

    for (const name of ["office.action.inspect", "office.action.validate"]) {
      const schema = parameters(name);
      expect(parseAndValidateToolArguments(JSON.stringify(valid), schema).error).toBeUndefined();
      expect(
        parseAndValidateToolArguments(
          JSON.stringify({ ...valid, params: { ...valid.params, paperSize: "a4" } }),
          schema,
        ).error,
      ).toContain("paperSize");
      expect(
        parseAndValidateToolArguments(
          JSON.stringify({ ...valid, params: { ...valid.params, host: "word" } }),
          schema,
        ).error,
      ).toContain("host");
    }
  });
});
