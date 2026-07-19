import { CHART_TYPES } from "../host/chartTypes";
import type { ToolDefinition } from "./types";

export const CHART_SERIES_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "chart.series.list",
    description:
      "列出指定图表的 series（index 1-based；name/chartType/smooth）。仅 Office.js；WPS unsupported",
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string", minLength: 1 },
        chartName: { type: "string", minLength: 1 },
      },
      required: ["sheetName", "chartName"],
      additionalProperties: false,
    },
  },
  {
    name: "chart.series.update",
    description:
      "更新指定 series 浅层属性：newName/chartType/smooth（≥1 字段；seriesIndex 1-based；十种 chartType；series.chartType ExcelApi 1.7）。Office.js；WPS unsupported",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string", minLength: 1 },
        chartName: { type: "string", minLength: 1 },
        seriesIndex: { type: "integer", exclusiveMinimum: 0 },
        newName: { type: "string", minLength: 1 },
        chartType: { type: "string", enum: [...CHART_TYPES] },
        smooth: { type: "boolean" },
      },
      required: ["sheetName", "chartName", "seriesIndex"],
      additionalProperties: false,
    },
  },
];
