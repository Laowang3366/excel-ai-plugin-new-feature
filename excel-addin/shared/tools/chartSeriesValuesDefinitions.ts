import type { ToolDefinition } from "./types";

export const CHART_SERIES_VALUES_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "chart.series.values.update",
    description:
      "绑定图表 series values/xValues 同表 A1 源（ExcelApi 1.15 真源回读；dataBound=true）。WPS unsupported",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string", minLength: 1 },
        chartName: { type: "string", minLength: 1 },
        seriesIndex: { type: "integer", minimum: 1 },
        valuesRange: { type: "string", minLength: 1 },
        xValuesRange: { type: "string", minLength: 1 },
      },
      required: ["sheetName", "chartName", "seriesIndex"],
      additionalProperties: false,
    },
  },
];
