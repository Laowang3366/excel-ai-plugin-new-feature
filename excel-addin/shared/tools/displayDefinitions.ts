import type { ToolDefinition } from "./types";

export const DISPLAY_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "sheet.display.get",
    description:
      "读取工作表显示属性 tabColor/showGridlines/showHeadings（Office.js；WPS unsupported）",
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: { sheetName: { type: "string" } },
      required: ["sheetName"],
      additionalProperties: false,
    },
  },
  {
    name: "sheet.display.set",
    description:
      "设置工作表显示属性。tabColor 空串=自动色或 #RRGGBB；需至少一个更新字段（Office.js；WPS unsupported）",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        tabColor: { type: "string" },
        showGridlines: { type: "boolean" },
        showHeadings: { type: "boolean" },
      },
      required: ["sheetName"],
      additionalProperties: false,
    },
  },
];
