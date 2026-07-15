import { describe, expect, it } from "vitest";

import { ALL_TOOL_DEFINITIONS } from "./toolDefinitions";
import { parseAndValidateToolArguments } from "./toolSchema";

function parameters(name: string) {
  const definition = ALL_TOOL_DEFINITIONS.find((tool) => tool.name === name);
  if (!definition) throw new Error(`missing tool definition: ${name}`);
  return definition.parameters;
}

function apply(params: Record<string, unknown>) {
  return {
    app: "word",
    action: "style",
    operation: "formatLongDocument",
    filePath: "C:/documents/report.docx",
    params,
  };
}

describe("Word long-document formatting parameter schemas", () => {
  it("accepts formatting inspection validation", () => {
    const args = {
      app: "word",
      operation: "inspectDocumentFormatting",
      filePath: "C:/documents/report.docx",
      params: { host: "word", countPath: "paragraphCount", minCount: 1 },
    };

    for (const tool of ["office.action.inspect", "office.action.validate"]) {
      expect(
        parseAndValidateToolArguments(JSON.stringify(args), parameters(tool)).error,
      ).toBeUndefined();
    }
  });

  it("accepts explicit automatic heading detection and real layout fields", () => {
    const args = apply({
      host: "word",
      autoDetectHeadings: true,
      fontName: "微软雅黑",
      fontSize: 11,
      headerColor: "#D9EAF7",
      margins: { top: 2.54, bottom: 2.54, left: 3.17, right: 3.17 },
      headerFooter: { header: "Quarterly report", footer: "Confidential" },
      pageNumbers: true,
      toc: "create",
      position: "start",
      upperHeadingLevel: 1,
      lowerHeadingLevel: 3,
    });

    expect(
      parseAndValidateToolArguments(JSON.stringify(args), parameters("office.action.apply")).error,
    ).toBeUndefined();
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({ steps: [args] }),
        parameters("office.workflow.run"),
      ).error,
    ).toBeUndefined();
  });

  it("accepts explicit prefix and bounded-regex heading strategies", () => {
    const schema = parameters("office.action.apply");
    const cases = [
      { autoDetectHeadings: false, startsWith: "Chapter ", level: 1 },
      { autoDetectHeadings: false, pattern: "^\\d+\\.\\d+", level: 2 },
      {
        autoDetectHeadings: false,
        startsWith: "Appendix ",
        pattern: "^附录[一二三四五六七八九十]",
        level: 1,
      },
    ];

    for (const params of cases) {
      expect(
        parseAndValidateToolArguments(JSON.stringify(apply(params)), schema).error,
      ).toBeUndefined();
    }
  });

  it("rejects implicit all-paragraph heading formatting and fictional style objects", () => {
    const schema = parameters("office.action.apply");
    const invalid = [
      {},
      { autoDetectHeadings: false },
      { startsWith: "Chapter ", level: 1 },
      { autoDetectHeadings: true, startsWith: "Chapter " },
      { autoDetectHeadings: false, startsWith: "Chapter ", headingStyles: [{ level: 1 }] },
      { autoDetectHeadings: true, normalStyle: { fontName: "Arial" } },
      { autoDetectHeadings: true, sectionBreaks: [{ position: 100, type: "nextPage" }] },
      { autoDetectHeadings: true, orientation: "landscape" },
      { autoDetectHeadings: true, margins: { top: 2.54, gutter: 1 } },
      { autoDetectHeadings: true, headerFooter: { header: "Report", differentFirstPage: true } },
    ];

    for (const params of invalid) {
      expect(
        parseAndValidateToolArguments(JSON.stringify(apply(params)), schema).error,
      ).toBeDefined();
    }
  });

  it("bounds model-provided heading regular expressions", () => {
    const schema = parameters("office.action.apply");
    expect(
      parseAndValidateToolArguments(
        JSON.stringify(apply({ autoDetectHeadings: false, pattern: "a".repeat(513) })),
        schema,
      ).error,
    ).toBeDefined();

    const headingArgs = {
      app: "word",
      action: "style",
      operation: "applyHeadingStyles",
      filePath: "C:/documents/report.docx",
      params: { pattern: "a".repeat(513), level: 1 },
    };
    expect(parseAndValidateToolArguments(JSON.stringify(headingArgs), schema).error).toContain(
      "512",
    );
  });
});
