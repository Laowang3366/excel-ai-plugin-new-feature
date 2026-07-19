import type { ToolDefinition } from "./types";

export const FREEZE_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "sheet.freeze.get",
    description:
      "读取工作表冻结窗格位置（address/rowCount/columnCount；无冻结 null/0/0）。Office.js；WPS unsupported",
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: { sheetName: { type: "string" } },
      required: ["sheetName"],
      additionalProperties: false,
    },
  },
  {
    name: "sheet.freeze.set",
    description:
      "设置冻结：command=rows|columns|at|clear；rows/columns 需正整数 count；at 需 address；clear 无额外字段。Office.js；WPS unsupported",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        command: { type: "string", enum: ["rows", "columns", "at", "clear"] },
        count: { type: "integer", minimum: 1 },
        address: { type: "string" },
      },
      required: ["sheetName", "command"],
      additionalProperties: false,
    },
  },
];
