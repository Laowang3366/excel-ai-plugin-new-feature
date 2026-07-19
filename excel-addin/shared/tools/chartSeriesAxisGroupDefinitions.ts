import type { ToolDefinition } from "./types";

export const CHART_SERIES_AXIS_GROUP_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "chart.series.axisGroup.update",
    description:
      "更新指定 series 的 axisGroup：primary|secondary（seriesIndex 1-based）。Office.js；WPS unsupported",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string", minLength: 1 },
        chartName: { type: "string", minLength: 1 },
        seriesIndex: { type: "integer", exclusiveMinimum: 0 },
        axisGroup: { type: "string", enum: ["primary", "secondary"] },
      },
      required: ["sheetName", "chartName", "seriesIndex", "axisGroup"],
      additionalProperties: false,
    },
  },
];
