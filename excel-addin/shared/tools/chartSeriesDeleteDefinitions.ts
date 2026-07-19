import type { ToolDefinition } from "./types";

export const CHART_SERIES_DELETE_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "chart.series.delete",
    description:
      "删除指定图表 series（seriesIndex 1-based；回读剩余 series）。Office.js；WPS unsupported",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string", minLength: 1 },
        chartName: { type: "string", minLength: 1 },
        seriesIndex: { type: "integer", minimum: 1 },
      },
      required: ["sheetName", "chartName", "seriesIndex"],
      additionalProperties: false,
    },
  },
];
