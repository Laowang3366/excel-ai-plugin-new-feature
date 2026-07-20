import type { ToolDefinition } from "./types";

export const CHART_SOURCE_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "chart.source.update",
    description:
      "替换图表数据源：sourceRange 为同表 A1，或同工作簿跨表 Sheet2!A1:B10 / 'Sheet 2'!A1:B10；seriesBy auto|rows|columns（默认 auto）；回读 series 快照。拒绝外部工作簿/3D/多区域/结构化引用。Office.js Chart.setData(Range)；WPS unsupported",
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
