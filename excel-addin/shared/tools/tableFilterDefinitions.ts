import { TABLE_FILTER_ON } from "../host/tableFilterTypes";
import type { ToolDefinition } from "./types";

export const TABLE_FILTER_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "table.filter.get",
    description:
      "读取表格 AutoFilter 是否启用（ExcelApi 1.9 Table.autoFilter.enabled）。不返回完整 criteria 明细。WPS unsupported",
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string", minLength: 1 },
        tableName: { type: "string", minLength: 1 },
      },
      required: ["sheetName", "tableName"],
      additionalProperties: false,
    },
  },
  {
    name: "table.filter.apply",
    description:
      "对表格列应用 AutoFilter（ExcelApi 1.2）。columnIndex 为表内 1-based。filterOn: values|custom|topItems|bottomItems|topPercent|bottomPercent。values 模式需 values[]；custom 需 criterion1（可选 criterion2/operator and|or）；top/bottom 需正数 threshold。cellColor/fontColor/icon/dynamic → unsupported。WPS unsupported",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string", minLength: 1 },
        tableName: { type: "string", minLength: 1 },
        columnIndex: { type: "integer", minimum: 1 },
        filterOn: { type: "string", enum: [...TABLE_FILTER_ON] },
        values: {
          type: "array",
          items: { type: "string" },
          minItems: 1,
          maxItems: 256,
        },
        criterion1: { type: "string", minLength: 1 },
        criterion2: { type: "string", minLength: 1 },
        operator: { type: "string", enum: ["and", "or"] },
        threshold: { type: "number", exclusiveMinimum: 0 },
      },
      required: ["sheetName", "tableName", "columnIndex", "filterOn"],
      additionalProperties: false,
    },
  },
  {
    name: "table.filter.clear",
    description:
      "清除表格 AutoFilter 条件（ExcelApi 1.2 Table.autoFilter.clearCriteria）。WPS unsupported",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string", minLength: 1 },
        tableName: { type: "string", minLength: 1 },
      },
      required: ["sheetName", "tableName"],
      additionalProperties: false,
    },
  },
];
