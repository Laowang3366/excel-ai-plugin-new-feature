import type { ToolDefinition } from "./types";

export const TEMPLATE_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "workbook.template.apply",
    description:
      "Apply desktop-parity professional workbook template (Office.js ExcelApi 1.8). Presets professional|financial|dashboard|minimal paint UsedRange font + header row (bold/fill/center/wrap/rowHeight=24), optional autofit, showGridlines, freezeRows. Empty sheets skipped. Write→sync→load→sync host readback; autoFitVerified=false. WPS typed unsupported. Not real sideload verified.",
    riskLevel: "dangerous",
    parameters: {
      type: "object",
      additionalProperties: false,
      properties: {
        preset: {
          type: "string",
          enum: ["professional", "financial", "dashboard", "minimal"],
        },
        sheetNames: {
          type: "array",
          maxItems: 500,
          items: { type: "string", minLength: 1, maxLength: 255 },
        },
        allSheets: { type: "boolean" },
        fontName: { type: "string", minLength: 1, maxLength: 255 },
        fontSize: { type: "number", minimum: 1, maximum: 409 },
        autoFit: { type: "boolean" },
        showGridlines: { type: "boolean" },
        freezeRows: { type: "integer", minimum: 0, maximum: 1_048_576 },
      },
    },
  },
  {
    name: "workbook.template.capture",
    description:
      "Shallow capture of workbook formatting (desktop captureWorkbookTemplate/inspectWorkbookFormatting parity). Office.js ExcelApi 1.9. Returns template v1 with per-sheet usedRange/baseStyle/headerStyle/print snapshot. Not a full theme/CF/DV dump and not replayable. WPS typed unsupported. Not real sideload verified.",
    riskLevel: "safe",
    parameters: {
      type: "object",
      additionalProperties: false,
      properties: {},
    },
  },
];
