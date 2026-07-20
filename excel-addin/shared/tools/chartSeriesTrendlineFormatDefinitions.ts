import { CHART_LINE_STYLES } from "../host/chartSeriesTrendlineFormatTypes";
import type { ToolDefinition } from "./types";

export const CHART_SERIES_TRENDLINE_FORMAT_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "chart.series.trendlines.format.update",
    description:
      "更新图表 series 趋势线线型格式：color(#RRGGBB)/lineStyle/weight(pt)；seriesIndex+trendlineIndex 1-based；ExcelApi 1.7 format.line 写后回读；WPS unsupported",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string", minLength: 1 },
        chartName: { type: "string", minLength: 1 },
        seriesIndex: { type: "integer", exclusiveMinimum: 0 },
        trendlineIndex: { type: "integer", exclusiveMinimum: 0 },
        color: { type: "string", minLength: 6, maxLength: 7 },
        lineStyle: { type: "string", enum: [...CHART_LINE_STYLES] },
        weight: { type: "number", exclusiveMinimum: 0 },
      },
      required: ["sheetName", "chartName", "seriesIndex", "trendlineIndex"],
      additionalProperties: false,
    },
  },
];
