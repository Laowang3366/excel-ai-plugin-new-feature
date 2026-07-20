import { CHART_MARKER_STYLES } from "../host/chartSeriesMarkersTypes";
import type { ToolDefinition } from "./types";

export const CHART_SERIES_MARKERS_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "chart.series.markers.update",
    description:
      "更新图表 series 标记：markerStyle/markerSize(2-72)/markerBackgroundColor|markerForegroundColor(#RRGGBB)；seriesIndex 1-based；ExcelApi 1.7 写后回读；WPS unsupported",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string", minLength: 1 },
        chartName: { type: "string", minLength: 1 },
        seriesIndex: { type: "integer", exclusiveMinimum: 0 },
        markerStyle: { type: "string", enum: [...CHART_MARKER_STYLES] },
        markerSize: { type: "integer", minimum: 2, maximum: 72 },
        markerBackgroundColor: { type: "string", minLength: 7, maxLength: 7 },
        markerForegroundColor: { type: "string", minLength: 7, maxLength: 7 },
      },
      required: ["sheetName", "chartName", "seriesIndex"],
      additionalProperties: false,
    },
  },
];
