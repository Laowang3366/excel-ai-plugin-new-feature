import type { ToolDefinition } from "./types";

export const CHART_IMAGE_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "chart.image.get",
    description:
      "读取图表图像为内存 Base64（ExcelApi 1.2 Chart.getImage；不写路径/PDF）。WPS unsupported",
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string", minLength: 1 },
        chartName: { type: "string", minLength: 1 },
        width: { type: "integer", minimum: 1, maximum: 4096 },
        height: { type: "integer", minimum: 1, maximum: 4096 },
      },
      required: ["sheetName", "chartName"],
      additionalProperties: false,
    },
  },
];
