import type { ToolDefinition } from "./types";

export const CHART_SERIES_ADD_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "chart.series.add",
    description:
      "创建空图表 series（可选 name；未绑定 values/xValues，dataBound=false，图中不可见）。Office.js；WPS unsupported",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string", minLength: 1 },
        chartName: { type: "string", minLength: 1 },
        name: { type: "string", minLength: 1 },
      },
      required: ["sheetName", "chartName"],
      additionalProperties: false,
    },
  },
];
