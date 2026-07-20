import type { ToolDefinition } from "./types";

const TRENDLINE_TYPE_ENUM = [
  "linear",
  "exponential",
  "logarithmic",
  "movingAverage",
  "polynomial",
  "power",
] as const;

const commonProps = {
  sheetName: { type: "string", minLength: 1 },
  chartName: { type: "string", minLength: 1 },
  seriesIndex: { type: "integer", minimum: 1 },
} as const;

const fieldProps = {
  type: { type: "string", enum: [...TRENDLINE_TYPE_ENUM] },
  name: { type: "string", minLength: 1 },
  intercept: {
    description: "Finite number, or empty string for automatic intercept (Office.js write)",
  },
  polynomialOrder: { type: "integer", minimum: 2, maximum: 6 },
  movingAveragePeriod: { type: "integer", minimum: 2 },
  forwardPeriod: { type: "number", minimum: 0 },
  backwardPeriod: { type: "number", minimum: 0 },
  showEquation: { type: "boolean" },
  showRSquared: { type: "boolean" },
} as const;

export const CHART_SERIES_TRENDLINE_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "chart.series.trendlines.list",
    description:
      "列出图表 series 的趋势线（seriesIndex 1-based；trendlineIndex 1-based 插入序）。ExcelApi 1.7；1.8 字段 period/equation 在宿主支持时回读。WPS unsupported",
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: { ...commonProps },
      required: ["sheetName", "chartName", "seriesIndex"],
      additionalProperties: false,
    },
  },
  {
    name: "chart.series.trendlines.add",
    description:
      "为图表 series 添加趋势线（type 必填；可选 name/intercept/polynomialOrder/movingAveragePeriod；1.8：forward/backwardPeriod/showEquation/showRSquared）。写后宿主回读。WPS unsupported",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: { ...commonProps, ...fieldProps },
      required: ["sheetName", "chartName", "seriesIndex", "type"],
      additionalProperties: false,
    },
  },
  {
    name: "chart.series.trendlines.update",
    description:
      "更新指定趋势线（trendlineIndex 1-based；≥1 更新字段）。写后宿主回读。ExcelApi 1.7/1.8。WPS unsupported",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        ...commonProps,
        trendlineIndex: { type: "integer", minimum: 1 },
        ...fieldProps,
      },
      required: ["sheetName", "chartName", "seriesIndex", "trendlineIndex"],
      additionalProperties: false,
    },
  },
  {
    name: "chart.series.trendlines.delete",
    description:
      "删除指定趋势线（trendlineIndex 1-based；回读 remainingTrendlines）。ExcelApi 1.7。WPS unsupported",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        ...commonProps,
        trendlineIndex: { type: "integer", minimum: 1 },
      },
      required: ["sheetName", "chartName", "seriesIndex", "trendlineIndex"],
      additionalProperties: false,
    },
  },
];
