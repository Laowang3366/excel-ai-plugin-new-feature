import type { ToolDefinition } from "./types";

export const CHART_SOURCE_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "chart.source.update",
    description:
      "替换图表数据源：同表 A1 sourceRange + seriesBy auto|rows|columns（默认 auto）；回读 series 快照。Office.js；WPS unsupported",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string", minLength: 1 },
        chartName: { type: "string", minLength: 1 },
        sourceRange: { type: "string", minLength: 1 },
        seriesBy: { type: "string", enum: ["auto", "rows", "columns"] },
      },
      required: ["sheetName", "chartName", "sourceRange"],
      additionalProperties: false,
    },
  },
];
