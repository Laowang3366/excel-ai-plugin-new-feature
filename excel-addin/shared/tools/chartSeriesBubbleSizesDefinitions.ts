import type { ToolDefinition } from "./types";

export const CHART_SERIES_BUBBLE_SIZES_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "chart.series.bubbleSizes.update",
    description:
      "绑定图表 series bubble sizes 同表 A1 源（ExcelApi 1.15 getDimensionDataSourceString 真源回读；仅 bubble chart；dataBound=true）。WPS unsupported",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string", minLength: 1 },
        chartName: { type: "string", minLength: 1 },
        seriesIndex: { type: "integer", minimum: 1 },
        bubbleSizesRange: { type: "string", minLength: 1 },
      },
      required: ["sheetName", "chartName", "seriesIndex", "bubbleSizesRange"],
      additionalProperties: false,
    },
  },
];
