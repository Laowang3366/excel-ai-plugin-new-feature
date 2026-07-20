import { describe, expect, it } from "vitest";
import { TOOL_DEFINITIONS } from "../shared/tools";
import type { ToolName } from "../shared/tools";

const LEGACY_STRICT_TOOLS = [
  "host.status",
  "selection.get",
  "workbook.inspect",
  "sheet.list",
  "range.read",
  "range.write",
  "range.clear",
  "range.format.read",
  "range.format.write",
  "formula.read",
  "formula.write",
  "formula.context",
  "sheet.operation",
  "sheet.add",
  "sheet.rename",
  "sheet.delete",
  "table.list",
  "table.create",
  "table.delete",
  "conditionalFormat.list",
  "conditionalFormat.add",
  "conditionalFormat.delete",
  "dataValidation.read",
  "dataValidation.write",
  "dataValidation.clear",
  "chart.list",
  "chart.delete",
] as const;

const FORMAT_KEYS = [
  "fontName",
  "fontSize",
  "fontBold",
  "fontColor",
  "fillColor",
  "numberFormat",
  "horizontalAlignment",
  "verticalAlignment",
  "wrapText",
] as const;

const FORMAT_TYPES: Record<(typeof FORMAT_KEYS)[number], string> = {
  fontName: "string",
  fontSize: "number",
  fontBold: "boolean",
  fontColor: "string",
  fillColor: "string",
  numberFormat: "string",
  horizontalAlignment: "string",
  verticalAlignment: "string",
  wrapText: "boolean",
};

describe("phase36 model-visible tool schemas are closed", () => {
  it("registers exactly 80 tools without name loss or duplicates", () => {
    expect(TOOL_DEFINITIONS).toHaveLength(80);
    const names = TOOL_DEFINITIONS.map((d) => d.name);
    expect(new Set(names).size).toBe(80);
    for (const name of LEGACY_STRICT_TOOLS) {
      expect(names).toContain(name);
    }
    // ToolName union remains usable for every definition
    const asToolNames = names as ToolName[];
    expect(asToolNames).toHaveLength(80);
  });

  it("every definition has object parameters with additionalProperties false", () => {
    for (const def of TOOL_DEFINITIONS) {
      const params = def.parameters as {
        type?: string;
        additionalProperties?: boolean;
      };
      expect(params.type, def.name).toBe("object");
      expect(params.additionalProperties, def.name).toBe(false);
    }
  });

  it("legacy 27 tools are present and top-level closed", () => {
    for (const name of LEGACY_STRICT_TOOLS) {
      const def = TOOL_DEFINITIONS.find((d) => d.name === name);
      expect(def, name).toBeTruthy();
      expect((def!.parameters as { additionalProperties?: boolean }).additionalProperties).toBe(
        false,
      );
    }
  });

  it("range.format.write.format is a closed 9-field object schema", () => {
    const def = TOOL_DEFINITIONS.find((d) => d.name === "range.format.write");
    expect(def).toBeTruthy();
    const format = (
      def!.parameters as {
        properties?: {
          format?: {
            type?: string;
            additionalProperties?: boolean;
            properties?: Record<string, { type?: string }>;
          };
        };
      }
    ).properties?.format;
    expect(format?.type).toBe("object");
    expect(format?.additionalProperties).toBe(false);
    const keys = Object.keys(format?.properties ?? {}).sort();
    expect(keys).toEqual([...FORMAT_KEYS].sort());
    for (const key of FORMAT_KEYS) {
      expect(format?.properties?.[key]?.type, key).toBe(FORMAT_TYPES[key]);
    }
  });

  it("conditionalFormat.add and dataValidation.write keep nested rule closed", () => {
    for (const name of ["conditionalFormat.add", "dataValidation.write"] as const) {
      const def = TOOL_DEFINITIONS.find((d) => d.name === name);
      const rule = (
        def!.parameters as {
          properties?: {
            rule?: { additionalProperties?: boolean; type?: string };
          };
        }
      ).properties?.rule;
      expect(rule?.type, name).toBe("object");
      expect(rule?.additionalProperties, name).toBe(false);
    }
  });

  it("chart.create remains top-level strict", () => {
    const def = TOOL_DEFINITIONS.find((d) => d.name === "chart.create");
    expect((def!.parameters as { additionalProperties?: boolean }).additionalProperties).toBe(
      false,
    );
  });
});
