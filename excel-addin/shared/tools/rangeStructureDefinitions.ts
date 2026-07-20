import type { ToolDefinition } from "./types";

const RANGE_PROPERTIES = {
  sheetName: { type: "string", minLength: 1 },
  range: { type: "string", minLength: 1 },
} as const;

export const RANGE_STRUCTURE_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "range.insert",
    description:
      "插入区域并移动现有单元格（shift: down|right；ExcelApi 1.1 Range.insert）。WPS unsupported",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        ...RANGE_PROPERTIES,
        shift: { type: "string", enum: ["down", "right"] },
      },
      required: ["sheetName", "range", "shift"],
      additionalProperties: false,
    },
  },
  {
    name: "range.delete",
    description:
      "删除区域并移动剩余单元格（shift: up|left；ExcelApi 1.1 Range.delete）。WPS unsupported",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        ...RANGE_PROPERTIES,
        shift: { type: "string", enum: ["up", "left"] },
      },
      required: ["sheetName", "range", "shift"],
      additionalProperties: false,
    },
  },
  {
    name: "range.autofit",
    description:
      "自动调整区域行高/列宽（direction: rows|columns|both；ExcelApi 1.2），并回读实际尺寸。WPS unsupported",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      properties: {
        ...RANGE_PROPERTIES,
        direction: { type: "string", enum: ["rows", "columns", "both"] },
      },
      required: ["sheetName", "range", "direction"],
      additionalProperties: false,
    },
  },
];
