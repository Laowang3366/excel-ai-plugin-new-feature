import type { ToolDefinition } from "./types";

export const TABLE_SORT_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "table.sort.get",
    description:
      "读取表格当前排序字段（ExcelApi 1.2 Table.sort.fields）。columnIndex 为表内 1-based 回读。WPS unsupported",
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
    name: "table.sort.apply",
    description:
      "按列排序表格（ExcelApi 1.2 Table.sort.apply）。fields[] 最多 3 级；columnIndex 表内 1-based；ascending 默认 true；可选 matchCase。仅 value 排序；颜色/图标排序 → unsupported。WPS unsupported",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string", minLength: 1 },
        tableName: { type: "string", minLength: 1 },
        fields: {
          type: "array",
          minItems: 1,
          maxItems: 3,
          items: {
            type: "object",
            properties: {
              columnIndex: { type: "integer", minimum: 1 },
              ascending: { type: "boolean" },
            },
            required: ["columnIndex"],
            additionalProperties: false,
          },
        },
        matchCase: { type: "boolean" },
      },
      required: ["sheetName", "tableName", "fields"],
      additionalProperties: false,
    },
  },
  {
    name: "table.sort.clear",
    description: "清除表格排序状态（ExcelApi 1.2 Table.sort.clear）。WPS unsupported",
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
