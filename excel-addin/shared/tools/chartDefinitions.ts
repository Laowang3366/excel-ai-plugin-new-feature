import { CHART_TYPES } from "../host/chartTypes";
import type { ToolDefinition } from "./types";

export const CHART_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "chart.list",
    description: "列出图表（可选按工作表过滤）",
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: { sheetName: { type: "string" } },
      required: [],
    },
  },
  {
    name: "chart.create",
    description:
      "从源区域创建图表（column|line|bar|area|pie|scatter|doughnut|bubble|radar|linemarkers；省略 chartType 默认 column；create 用 ChartCollection.add enum ExcelApi 1.1）",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        sourceRange: { type: "string" },
        chartType: { type: "string", enum: [...CHART_TYPES] },
        name: { type: "string" },
        title: { type: "string" },
        left: { type: "number" },
        top: { type: "number" },
        width: { type: "number" },
        height: { type: "number" },
      },
      required: ["sheetName", "sourceRange"],
      additionalProperties: false,
    },
  },
  {
    name: "chart.delete",
    description: "删除指定图表",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        chartName: { type: "string" },
      },
      required: ["sheetName", "chartName"],
    },
  },
];
