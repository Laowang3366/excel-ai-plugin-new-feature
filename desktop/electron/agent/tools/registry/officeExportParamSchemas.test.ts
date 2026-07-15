import { describe, expect, it } from "vitest";

import { ALL_TOOL_DEFINITIONS } from "./toolDefinitions";
import { parseAndValidateToolArguments } from "./toolSchema";

function parameters(name: string) {
  const definition = ALL_TOOL_DEFINITIONS.find((tool) => tool.name === name);
  if (!definition) throw new Error(`missing tool definition: ${name}`);
  return definition.parameters;
}

describe("Office PDF export parameter schemas", () => {
  it("accepts Excel combined and separate worksheet exports", () => {
    const apply = parameters("office.action.apply");
    const base = {
      app: "excel",
      action: "edit",
      operation: "exportSheetsToPdf",
      filePath: "C:/book.xlsx",
      outputPath: "C:/exports/book.pdf",
    };

    expect(
      parseAndValidateToolArguments(
        JSON.stringify({
          ...base,
          params: { sheetNames: ["Summary", "Details"], mode: "combined", overwrite: true },
        }),
        apply,
      ).error,
    ).toBeUndefined();
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({
          steps: [
            {
              ...base,
              params: {
                host: "wps",
                sheetNames: ["Summary"],
                mode: "separate",
                outputDirectory: "C:/exports/sheets",
              },
            },
          ],
        }),
        parameters("office.workflow.run"),
      ).error,
    ).toBeUndefined();
  });

  it("rejects malformed worksheet export parameters", () => {
    const base = {
      app: "excel",
      action: "edit",
      operation: "exportSheetsToPdf",
      filePath: "C:/book.xlsx",
    };
    const apply = parameters("office.action.apply");

    expect(
      parseAndValidateToolArguments(JSON.stringify({ ...base, params: { mode: "zip" } }), apply)
        .error,
    ).toContain("mode");
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({ ...base, params: { mode: "combined", shellCommand: "whoami" } }),
        apply,
      ).error,
    ).toContain("shellCommand");
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({ ...base, params: { sheetNames: [""] } }),
        apply,
      ).error,
    ).toContain("sheetNames");
  });

  it("keeps app-specific exportPdf params for Excel and Word", () => {
    const apply = parameters("office.action.apply");
    const excel = {
      app: "excel",
      action: "edit",
      operation: "exportPdf",
      filePath: "C:/book.xlsx",
      outputPath: "C:/book.pdf",
      params: { host: "excel", scope: "sheet" },
    };
    const word = {
      app: "word",
      action: "edit",
      operation: "exportPdf",
      filePath: "C:/report.docx",
      outputPath: "C:/report.pdf",
      params: { host: "word" },
    };

    expect(parseAndValidateToolArguments(JSON.stringify(excel), apply).error).toBeUndefined();
    expect(parseAndValidateToolArguments(JSON.stringify(word), apply).error).toBeUndefined();
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({ ...word, params: { host: "word", scope: "sheet" } }),
        apply,
      ).error,
    ).toContain("scope");
    expect(
      parseAndValidateToolArguments(JSON.stringify({ ...excel, params: { host: "word" } }), apply)
        .error,
    ).toContain("host");
  });
});
