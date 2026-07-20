import type { ToolDefinition } from "./types";

const CF_OPERATORS = [
  "greaterThan",
  "greaterThanOrEqualTo",
  "lessThan",
  "lessThanOrEqualTo",
  "equalTo",
  "notEqualTo",
  "between",
  "notBetween",
] as const;

const DV_OPERATORS = [
  "between",
  "notBetween",
  "equalTo",
  "notEqualTo",
  "greaterThan",
  "greaterThanOrEqualTo",
  "lessThan",
  "lessThanOrEqualTo",
] as const;

const DV_TYPES = [
  "list",
  "wholeNumber",
  "decimal",
  "date",
  "time",
  "textLength",
  "custom",
] as const;

export const CONDITIONAL_FORMAT_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "conditionalFormat.list",
    description:
      "列出区域条件格式。返回 hostType + kind(cellValue|custom|unsupported) + supported；不把 DataBar/ColorScale/IconSet/TopBottom/PresetCriteria/ContainsText 伪装为 cellValue。ExcelApi 1.6；WPS unsupported",
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        range: { type: "string" },
      },
      required: ["sheetName", "range"],
      additionalProperties: false,
    },
  },
  {
    name: "conditionalFormat.add",
    description:
      "添加条件格式（仅 cellValue/custom）。cellValue 支持完整比较运算符（含 >=/<=/!=）；custom 使用表达式 formula。颜色仅 #RRGGBB。写后从宿主集合回读。ExcelApi 1.6；WPS unsupported",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        range: { type: "string" },
        rule: {
          type: "object",
          properties: {
            kind: { type: "string", enum: ["cellValue", "custom"] },
            operator: { type: "string", enum: [...CF_OPERATORS] },
            formula1: { type: "string" },
            formula2: { type: "string" },
            formula: { type: "string" },
            fillColor: { type: "string" },
            fontColor: { type: "string" },
          },
          required: ["kind"],
          additionalProperties: false,
        },
      },
      required: ["sheetName", "range", "rule"],
      additionalProperties: false,
    },
  },
  {
    name: "conditionalFormat.delete",
    description: "按 id 删除区域条件格式；删除后回读确认。ExcelApi 1.6；WPS unsupported",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        range: { type: "string" },
        id: { type: "string" },
      },
      required: ["sheetName", "range", "id"],
      additionalProperties: false,
    },
  },
];

export const DATA_VALIDATION_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "dataValidation.read",
    description:
      "读取区域数据验证。支持 list/wholeNumber/decimal/date/time/textLength/custom；Inconsistent/MixedCriteria 诚实标记 limitations 且 rule=null。list 公式/区域源不拆成 listValues。ExcelApi 1.8；WPS unsupported",
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        range: { type: "string" },
      },
      required: ["sheetName", "range"],
      additionalProperties: false,
    },
  },
  {
    name: "dataValidation.write",
    description:
      "写入数据验证。list：listValues 内联 与 formula1 区域源互斥（区域源传 Range 代理）。比较型：完整 operator；between/notBetween 需 formula2。custom：formula1 为公式。errorAlert/prompt 本批不实现。写后宿主回读。ExcelApi 1.8；WPS unsupported",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        range: { type: "string" },
        rule: {
          type: "object",
          properties: {
            type: { type: "string", enum: [...DV_TYPES] },
            operator: { type: "string", enum: [...DV_OPERATORS] },
            formula1: { type: "string" },
            formula2: { type: "string" },
            listValues: {
              type: "array",
              items: { type: "string", minLength: 1 },
              minItems: 1,
            },
            allowBlank: { type: "boolean" },
          },
          required: ["type"],
          additionalProperties: false,
        },
      },
      required: ["sheetName", "range", "rule"],
      additionalProperties: false,
    },
  },
  {
    name: "dataValidation.clear",
    description: "清除区域数据验证；清除后回读确认。ExcelApi 1.8；WPS unsupported",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        range: { type: "string" },
      },
      required: ["sheetName", "range"],
      additionalProperties: false,
    },
  },
];
