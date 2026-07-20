import type { ToolDefinition } from "./types";

const PIVOT_FIELD_ITEM = {
  oneOf: [
    { type: "string", minLength: 1 },
    {
      type: "object",
      properties: {
        name: { type: "string", minLength: 1 },
        function: {
          type: "string",
          enum: ["sum", "count", "average", "max", "min"],
        },
        caption: { type: "string" },
      },
      required: ["name"],
      additionalProperties: false,
    },
  ],
} as const;

const PIVOT_FIELDS = {
  type: "array",
  maxItems: 64,
  items: PIVOT_FIELD_ITEM,
} as const;

export const PIVOT_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "pivot.list",
    description:
      "列出当前工作簿（或指定工作表）的数据透视表：name/sheet/source|destination（可得时）/row|column|filter|data hierarchy 摘要。Office.js ExcelApi 1.8；无法可靠取得的字段写入 limitations，不伪造。WPS typed unsupported。",
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string", minLength: 1 },
      },
      required: [],
      additionalProperties: false,
    },
  },
  {
    name: "pivot.create",
    description:
      "在当前工作簿创建数据透视表。必填 sourceSheetName+sourceAddress（同簿单区域 A1）；可选 name、destination（空→Pivots 表自动位置 A1 或现有透视下方）、rowFields/columnFields/filterFields/dataFields（≤64；字符串或 {name,function?,caption?}，function 仅 dataFields：sum|count|average|max|min）。必填 advancedIntent=interactive-pivot。写后 sync 回读校验名称与 hierarchy 数量。拒绝外部/3D/多区域/结构化源；refreshConnections 不适用。Office.js ExcelApi 1.8；WPS unsupported；无切片器。",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        advancedIntent: { type: "string", enum: ["interactive-pivot"] },
        sourceSheetName: { type: "string", minLength: 1 },
        sourceAddress: { type: "string", minLength: 1 },
        name: { type: "string", minLength: 1 },
        destination: { type: "string", minLength: 1 },
        rowFields: PIVOT_FIELDS,
        columnFields: PIVOT_FIELDS,
        filterFields: PIVOT_FIELDS,
        dataFields: PIVOT_FIELDS,
      },
      required: ["advancedIntent", "sourceSheetName", "sourceAddress"],
      additionalProperties: false,
    },
  },
  {
    name: "pivot.refresh",
    description:
      "刷新当前工作簿透视表：省略 sheetName/name 刷新全部；可按 sheet 或 name 精确刷新。写后回读确认对象仍在。refreshConnections=true 拒绝（无 Office.js Workbook.RefreshAll 合同）。必填 advancedIntent=interactive-pivot。Office.js ExcelApi 1.8；WPS unsupported。",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        advancedIntent: { type: "string", enum: ["interactive-pivot"] },
        sheetName: { type: "string", minLength: 1 },
        name: { type: "string", minLength: 1 },
        refreshConnections: { type: "boolean" },
      },
      required: ["advancedIntent"],
      additionalProperties: false,
    },
  },
];
