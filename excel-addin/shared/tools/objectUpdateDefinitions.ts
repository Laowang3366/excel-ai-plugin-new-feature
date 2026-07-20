import { CHART_TYPES } from "../host/chartTypes";
import type { ToolDefinition } from "./types";

export const OBJECT_UPDATE_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "table.update",
    description:
      "更新表格属性或范围：newName/style/showHeaders/showTotals/showFilterButton/showBandedRows/showBandedColumns/showFirstColumn/showLastColumn/resizeAddress（至少一项；首末列高亮 ExcelApi 1.3；resizeAddress 为同表单区域 A1；Office.js；WPS unsupported）",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        tableName: { type: "string" },
        newName: { type: "string", minLength: 1 },
        style: { type: "string", minLength: 1 },
        showHeaders: { type: "boolean" },
        showTotals: { type: "boolean" },
        showFilterButton: { type: "boolean" },
        showBandedRows: { type: "boolean" },
        showBandedColumns: { type: "boolean" },
        showFirstColumn: { type: "boolean" },
        showLastColumn: { type: "boolean" },
        resizeAddress: { type: "string", minLength: 1 },
      },
      required: ["sheetName", "tableName"],
      additionalProperties: false,
    },
  },
  {
    name: "chart.update",
    description:
      "更新图表浅层属性：newName/chartType/title/showTitle/style/showLegend/left/top/width/height（至少一项；chartType 十种；Chart.chartType ExcelApi 1.7；Office.js；WPS unsupported）",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        chartName: { type: "string" },
        newName: { type: "string", minLength: 1 },
        chartType: {
          type: "string",
          enum: [...CHART_TYPES],
        },
        title: { type: "string" },
        showTitle: { type: "boolean" },
        style: { type: "integer", exclusiveMinimum: 0 },
        showLegend: { type: "boolean" },
        left: { type: "number" },
        top: { type: "number" },
        width: { type: "number", exclusiveMinimum: 0 },
        height: { type: "number", exclusiveMinimum: 0 },
      },
      required: ["sheetName", "chartName"],
      additionalProperties: false,
    },
  },
];
