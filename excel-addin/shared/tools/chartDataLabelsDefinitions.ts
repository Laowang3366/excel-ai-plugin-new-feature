import type { ToolDefinition } from "./types";
import { CHART_DATA_LABEL_POSITIONS } from "../host/chartDataLabelsTypes";

export const CHART_DATA_LABELS_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "chart.series.dataLabels.update",
    description:
      "更新指定 series 数据标签：enabled-only→ExcelApi 1.7 hasDataLabels（不访问 dataLabels）；showValue/showCategoryName/showSeriesName/numberFormat/showPercentage/showBubbleSize/showLegendKey/separator/position→ExcelApi 1.8 完整快照（≥1 字段；seriesIndex 1-based；enabled=false 不可与其它标签字段同传；position 不含 Invalid；separator 原样含空串）。Office.js 真回读；WPS unsupported",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string", minLength: 1 },
        chartName: { type: "string", minLength: 1 },
        seriesIndex: { type: "integer", exclusiveMinimum: 0 },
        enabled: { type: "boolean" },
        showValue: { type: "boolean" },
        showCategoryName: { type: "boolean" },
        showSeriesName: { type: "boolean" },
        numberFormat: { type: "string", minLength: 1 },
        showPercentage: { type: "boolean" },
        showBubbleSize: { type: "boolean" },
        showLegendKey: { type: "boolean" },
        separator: { type: "string" },
        position: {
          type: "string",
          enum: [...CHART_DATA_LABEL_POSITIONS],
        },
      },
      required: ["sheetName", "chartName", "seriesIndex"],
      additionalProperties: false,
    },
  },
];
