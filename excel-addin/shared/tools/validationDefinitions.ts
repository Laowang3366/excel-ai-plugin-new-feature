import type { ToolDefinition } from "./types";

export const CONDITIONAL_FORMAT_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "conditionalFormat.list",
    description: "列出区域条件格式（Office.js；WPS unsupported）",
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
      "添加条件格式。rule.kind=cellValue|custom；cellValue 需 operator/formula1；custom 需 formula",
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
            operator: {
              type: "string",
              enum: ["greaterThan", "lessThan", "equalTo", "between", "notBetween"],
            },
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
    description: "按 id 删除区域条件格式",
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
    description: "读取区域数据验证规则（Office.js；WPS unsupported）",
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
      "写入数据验证。rule.type=list|wholeNumber；list 用 listValues；wholeNumber 用 operator/formula1/formula2",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        range: { type: "string" },
        rule: {
          type: "object",
          properties: {
            type: { type: "string", enum: ["list", "wholeNumber"] },
            operator: {
              type: "string",
              enum: ["between", "notBetween", "equalTo", "greaterThan", "lessThan"],
            },
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
    description: "清除区域数据验证",
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
