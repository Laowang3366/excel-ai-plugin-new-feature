import type { ToolDefinition } from "./types";

const GEOMETRIC_ENUM = [
  "rectangle",
  "ellipse",
  "triangle",
  "diamond",
  "rightArrow",
] as const;

const NON_EMPTY = { type: "string", minLength: 1 } as const;
const POSITIVE = { type: "number", exclusiveMinimum: 0 } as const;

export const SHAPE_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "shape.list",
    description: "列出工作表形状（可选按 sheetName 过滤）",
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: { sheetName: { type: "string", minLength: 1 } },
      required: [],
      additionalProperties: false,
    },
  },
  {
    name: "shape.create",
    description:
      "创建形状：kind=geometric（geometricType: rectangle|ellipse|triangle|diamond|rightArrow）或 kind=textBox（可选 text）。浅层 name/left/top/width/height。",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: NON_EMPTY,
        kind: { type: "string", enum: ["geometric", "textBox"] },
        geometricType: { type: "string", enum: [...GEOMETRIC_ENUM] },
        text: { type: "string" },
        name: NON_EMPTY,
        left: { type: "number" },
        top: { type: "number" },
        width: POSITIVE,
        height: POSITIVE,
      },
      required: ["sheetName", "kind"],
      additionalProperties: false,
    },
  },
  {
    name: "shape.delete",
    description: "删除指定形状",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: NON_EMPTY,
        shapeName: NON_EMPTY,
      },
      required: ["sheetName", "shapeName"],
      additionalProperties: false,
    },
  },
  {
    name: "shape.update",
    description:
      "浅层更新形状：newName/left/top/width/height/text/visible（≥1 字段；w/h>0）。无 fill/line/rotation/zOrder。",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        sheetName: NON_EMPTY,
        shapeName: NON_EMPTY,
        newName: NON_EMPTY,
        left: { type: "number" },
        top: { type: "number" },
        width: POSITIVE,
        height: POSITIVE,
        text: { type: "string" },
        visible: { type: "boolean" },
      },
      required: ["sheetName", "shapeName"],
      additionalProperties: false,
    },
  },
];
