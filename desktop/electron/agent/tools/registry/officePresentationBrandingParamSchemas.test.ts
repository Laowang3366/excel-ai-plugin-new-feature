import { describe, expect, it } from "vitest";

import { ALL_TOOL_DEFINITIONS } from "./toolDefinitions";
import { parseAndValidateToolArguments } from "./toolSchema";

function parameters(name: string) {
  const definition = ALL_TOOL_DEFINITIONS.find((tool) => tool.name === name);
  if (!definition) throw new Error(`missing tool definition: ${name}`);
  return definition.parameters;
}

function apply(operation: string, params: Record<string, unknown>) {
  return {
    app: "presentation",
    action: "style",
    operation,
    filePath: "C:/presentations/report.pptx",
    target: "slide:2",
    params,
  };
}

describe("PowerPoint branding and layout parameter schemas", () => {
  it("accepts real branding fields with an explicit slide-number decision", () => {
    const args = apply("applyMasterBranding", {
      host: "powerpoint",
      showSlideNumber: true,
      templatePath: "C:/templates/corporate.potx",
      backgroundColor: "#FFFFFF",
      accentColor: "1F4E79",
      themeColors: [
        { index: 1, value: "#FFFFFF" },
        { index: 5, value: "#1F4E79" },
      ],
      fontName: "Aptos",
      fontMap: { Arial: "Aptos", 宋体: "微软雅黑" },
      applyAccentToText: false,
      logoPath: "C:/assets/logo.png",
      logoWidth: 100,
      logoTop: 18,
      footerText: "Confidential",
      layoutMap: [
        { slideIndex: 1, layoutName: "Title Slide" },
        { slideName: "Summary", layoutName: "Title and Content" },
      ],
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

  it("rejects implicit branding defaults and fictional nested fields", () => {
    const schema = parameters("office.action.apply");
    const invalid = [
      {},
      { fontName: "Aptos" },
      { showSlideNumber: true, themeColors: [{ index: 13, value: "#FFFFFF" }] },
      { showSlideNumber: true, themeColors: [{ index: 1, value: "red" }] },
      { showSlideNumber: true, fontMap: { Arial: "" } },
      {
        showSlideNumber: true,
        layoutMap: [{ slideIndex: 1, slideName: "Intro", layoutName: "Title" }],
      },
      { showSlideNumber: true, themeName: "Corporate" },
    ];

    for (const params of invalid) {
      expect(
        parseAndValidateToolArguments(JSON.stringify(apply("applyMasterBranding", params)), schema)
          .error,
      ).toBeDefined();
    }
  });

  it("accepts precise shape, table, chart, and crop edits", () => {
    const args = apply("layoutElements", {
      mode: "precise",
      edits: [
        {
          shapeName: "Title 1",
          left: 40,
          top: 20,
          width: 600,
          text: "Quarterly results",
          fontName: "Aptos Display",
          fontSize: 28,
        },
        {
          shapeIndex: 3,
          tableCells: [{ row: 1, column: 1, text: "Metric", fillColor: "#D9EAF7" }],
          chart: { chartType: 51, title: "Revenue", hasLegend: false },
          crop: { left: 0, right: 4, top: 0, bottom: 4 },
        },
      ],
    });

    expect(
      parseAndValidateToolArguments(JSON.stringify(args), parameters("office.action.apply")).error,
    ).toBeUndefined();
  });

  it("rejects implicit whole-slide grids and ambiguous edit selectors", () => {
    const schema = parameters("office.action.apply");
    const invalid = [
      {},
      { mode: "grid" },
      { mode: "precise", edits: [] },
      { mode: "precise", edits: [{ left: 10 }] },
      { mode: "precise", edits: [{ shapeName: "Title 1", shapeIndex: 1, left: 10 }] },
      { mode: "precise", edits: [{ shapeName: "Title 1", opacity: 0.5 }] },
      { mode: "grid", shapeNames: ["Chart 1"], columns: 0 },
    ];

    for (const params of invalid) {
      expect(
        parseAndValidateToolArguments(JSON.stringify(apply("layoutElements", params)), schema)
          .error,
      ).toBeDefined();
    }
  });

  it("accepts explicit grid, alignment, distribution, and fit strategies", () => {
    const schema = parameters("office.action.apply");
    const valid = [
      {
        mode: "grid",
        shapeNames: ["Chart 1", "Chart 2"],
        columns: 2,
        margin: 40,
        gap: 16,
        rowHeight: 180,
        resize: true,
      },
      { mode: "align", shapeNames: ["TextBox 1", "TextBox 2"], align: "left" },
      {
        mode: "distribute",
        shapeNames: ["Chart 1", "Chart 2", "Chart 3"],
        distribute: "horizontal",
      },
      { mode: "fit", shapeNames: ["Picture 1"], fitToSlide: true },
    ];

    for (const params of valid) {
      expect(
        parseAndValidateToolArguments(JSON.stringify(apply("layoutElements", params)), schema)
          .error,
      ).toBeUndefined();
    }
  });

  it("rejects cross-strategy and fictional layout fields", () => {
    const schema = parameters("office.action.apply");
    const invalid = [
      { mode: "align", shapeNames: ["TextBox 1"], distribute: "horizontal" },
      { mode: "distribute", shapeNames: ["Chart 1"], align: "left" },
      { mode: "fit", shapeNames: ["Picture 1"], fitToSlide: false },
      { mode: "grid", shapeNames: ["Chart 1"], align: "left" },
      { mode: "auto", shapeNames: ["Chart 1"], snapToGrid: true },
      { mode: "freeform", shapeNames: ["Chart 1"] },
    ];

    for (const params of invalid) {
      expect(
        parseAndValidateToolArguments(JSON.stringify(apply("layoutElements", params)), schema)
          .error,
      ).toBeDefined();
    }
  });
});
