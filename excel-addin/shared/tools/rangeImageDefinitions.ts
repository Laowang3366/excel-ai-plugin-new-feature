import type { ToolDefinition } from "./types";

export const RANGE_IMAGE_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "range.image.get",
    description:
      "读取工作表区域 PNG 为内存 Base64（ExcelApi 1.7 Range.getImage；不写路径/PDF/MIME）。WPS unsupported",
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string", minLength: 1 },
        range: { type: "string", minLength: 1 },
      },
      required: ["sheetName", "range"],
      additionalProperties: false,
    },
  },
];
