import { describe, expect, it } from "vitest";

import { ALL_TOOL_DEFINITIONS } from "./toolDefinitions";
import { parseAndValidateToolArguments } from "./toolSchema";

function parameters(name: string) {
  const definition = ALL_TOOL_DEFINITIONS.find((tool) => tool.name === name);
  if (!definition) throw new Error(`missing tool definition: ${name}`);
  return definition.parameters;
}

function action(
  app: "excel" | "word" | "presentation",
  operation: string,
  params: Record<string, unknown>,
) {
  return {
    app,
    action: operation.startsWith("inspect") ? "inspect" : "edit",
    operation,
    filePath:
      app === "excel"
        ? "C:/reports/source.xlsx"
        : `C:/reports/target.${app === "word" ? "docx" : "pptx"}`,
    params,
  };
}

describe("cross-Office parameter schemas", () => {
  it("accepts explicit Word and PowerPoint export branches", () => {
    const apply = parameters("office.action.apply");
    const word = action("excel", "exportRangeToWord", {
      sourceHost: "excel",
      wordHost: "word",
      instanceId: "word-window-1",
      linked: true,
      linkId: "sales-word",
      sourceType: "range",
      title: "Sales",
    });
    const presentation = action("excel", "exportRangeToPresentation", {
      sourceHost: "wps",
      presentationHost: "powerpoint",
      linked: true,
      sourceType: "chart",
      chartName: "RevenueChart",
      left: 40,
      top: 80,
      width: 640,
    });

    expect(parseAndValidateToolArguments(JSON.stringify(word), apply).error).toBeUndefined();
    expect(
      parseAndValidateToolArguments(JSON.stringify(presentation), apply).error,
    ).toBeUndefined();
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({ steps: [word, presentation] }),
        parameters("office.workflow.run"),
      ).error,
    ).toBeUndefined();
  });

  it("requires stable incremental identities and rejects cross-target fields", () => {
    const apply = parameters("office.action.apply");
    const invalid = [
      action("excel", "exportRangeToWord", { linked: true, updateExisting: true }),
      action("excel", "exportRangeToPresentation", {
        linked: true,
        updateExisting: true,
        linkId: "sales",
        overwrite: true,
      }),
      action("excel", "exportRangeToWord", { linked: true, left: 20 }),
      action("excel", "exportRangeToWord", { linked: true, asPicture: true }),
      action("excel", "exportRangeToPresentation", { linked: true, asPicture: true }),
      action("excel", "exportRangeToPresentation", { linked: true, sourceType: "chart" }),
      action("excel", "exportRangeToPresentation", { linked: true, width: 0 }),
    ];

    for (const args of invalid) {
      expect(parseAndValidateToolArguments(JSON.stringify(args), apply).error).toBeDefined();
    }
  });

  it("accepts strict report sections and requires every incremental linkId", () => {
    const apply = parameters("office.action.apply");
    const create = action("excel", "buildReportPackage", {
      linked: true,
      overwrite: true,
      outputDirectory: "C:/reports/package",
      baseName: "quarterly",
      sections: [
        { linkId: "sales", sheetName: "Sales", range: "A1:D20", title: "Sales" },
        {
          linkId: "trend",
          sheetName: "Dashboard",
          range: "A1:H30",
          sourceType: "chart",
          chartName: "TrendChart",
          width: 720,
        },
      ],
    });
    const update = action("excel", "buildReportPackage", {
      linked: true,
      updateExisting: true,
      wordOutputPath: "C:/reports/quarterly.docx",
      presentationOutputPath: "C:/reports/quarterly.pptx",
      sections: [{ linkId: "sales", range: "A1:D21" }],
    });

    expect(parseAndValidateToolArguments(JSON.stringify(create), apply).error).toBeUndefined();
    expect(parseAndValidateToolArguments(JSON.stringify(update), apply).error).toBeUndefined();

    const invalidParams = [
      { linked: true, sections: [] },
      { linked: false, sections: [{ range: "A1:B2" }] },
      { linked: true, sections: [{ title: "Missing range" }] },
      { linked: true, sections: [{ range: "A1:B2", sourceType: "chart" }] },
      { linked: true, sections: [{ range: "A1:B2", shellCommand: "whoami" }] },
      { linked: true, updateExisting: true, sections: [{ range: "A1:B2" }] },
    ];
    for (const params of invalidParams) {
      expect(
        parseAndValidateToolArguments(
          JSON.stringify(action("excel", "buildReportPackage", params)),
          apply,
        ).error,
      ).toBeDefined();
    }
  });

  it("accepts deterministic linked-content inspection, refresh and relink", () => {
    const inspect = parameters("office.action.inspect");
    const apply = parameters("office.action.apply");

    for (const app of ["word", "presentation"] as const) {
      expect(
        parseAndValidateToolArguments(
          JSON.stringify(
            action(app, "inspectLinkedOfficeContent", { linkId: "sales", minCount: 1 }),
          ),
          inspect,
        ).error,
      ).toBeUndefined();
      expect(
        parseAndValidateToolArguments(
          JSON.stringify(action(app, "refreshLinkedOfficeContent", { sourceHost: "excel" })),
          apply,
        ).error,
      ).toBeUndefined();
      expect(
        parseAndValidateToolArguments(
          JSON.stringify(
            action(app, "relinkLinkedOfficeContent", {
              linkId: "sales",
              sourcePath: "C:/reports/source-v2.xlsx",
            }),
          ),
          apply,
        ).error,
      ).toBeUndefined();
    }
  });

  it("rejects ambiguous relink aliases and malformed routing", () => {
    const apply = parameters("office.action.apply");
    const invalid = [
      action("word", "relinkLinkedOfficeContent", { sourcePath: "C:/reports/source.xlsx" }),
      action("presentation", "relinkLinkedOfficeContent", { linkId: "sales" }),
      action("word", "relinkLinkedOfficeContent", {
        linkId: "sales",
        newSourcePath: "C:/reports/source.xlsx",
      }),
      action("word", "refreshLinkedOfficeContent", { host: "powerpoint" }),
      action("presentation", "refreshLinkedOfficeContent", { sourceHost: "word" }),
    ];

    for (const args of invalid) {
      expect(parseAndValidateToolArguments(JSON.stringify(args), apply).error).toBeDefined();
    }
  });
});
