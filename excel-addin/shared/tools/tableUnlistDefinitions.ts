import type { ToolDefinition } from "./types";

export const TABLE_UNLIST_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "table.unlist",
    description:
      "将表格转为普通区域并保留全部数据（ExcelApi 1.2 Table.convertToRange）。不删除单元格。table.delete 仍为硬删除。WPS unsupported",
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
