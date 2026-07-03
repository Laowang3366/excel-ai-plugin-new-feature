/**
 * 单元格范围工具定义
 *
 * 包含范围读写、清理和选区读取工具。
 */

import type { ToolDefinition } from "../../shared/types";

/** 读取范围数据 */
const RANGE_READ_DEF: ToolDefinition = {
  name: "range.read",
  description: "【只读】读取单元格数据，默认返回二维数组。用途：查看数据、确认写入结果、分析内容。动态数组公式写入后验证溢出区时可传 expand:\"spill\"；读取连续结果区可传 expand:\"currentRegion\"。注意：此工具不能写入或修改任何数据，写入请用 range.write",
  parameters: {
    type: "object",
    properties: {
      sheetName: { type: "string", description: "工作表名称" },
      range: { type: "string", description: "单元格范围，如 A1:C10" },
      expand: {
        type: "string",
        enum: ["none", "spill", "currentArray", "currentRegion"],
        description: "可选扩展读取模式。none=只读指定范围；spill=从锚点读取动态数组溢出区；currentArray=读取传统数组区域；currentRegion=读取当前连续区域",
      },
    },
    required: ["sheetName", "range"],
  },
  riskLevel: "safe",
  requiresApproval: false,
};

/** 写入范围数据 */
const RANGE_WRITE_DEF: ToolDefinition = {
  name: "range.write",
  description: "【写入】向单元格写入数据或公式。用途：填入计算值、写入公式、批量更新单元格。参数 values 为二维数组（行×列），写入公式时以=开头。注意：此工具只能写入，不能读取，读取请用 range.read",
  parameters: {
    type: "object",
    properties: {
      sheetName: { type: "string", description: "工作表名称" },
      range: { type: "string", description: "起始单元格，如 A1" },
      values: {
        type: "array",
        description: "二维数组数据，外层对应行、内层对应列",
        items: { type: "array", items: {} },
      },
    },
    required: ["sheetName", "range", "values"],
  },
  riskLevel: "moderate",
  requiresApproval: true,
};

/** 清除范围数据 */
const RANGE_CLEAR_DEF: ToolDefinition = {
  name: "range.clear",
  description: "清空工作表中指定范围的单元格数据（值和格式）。用于清除旧数据、重置区域",
  parameters: {
    type: "object",
    properties: {
      sheetName: { type: "string", description: "工作表名称" },
      range: { type: "string", description: "单元格范围" },
    },
    required: ["sheetName", "range"],
  },
  riskLevel: "moderate",
  requiresApproval: true,
  isFileDeletion: true,
};

/** 获取选区信息 */
const SELECTION_GET_DEF: ToolDefinition = {
  name: "selection.get",
  description: "获取用户当前在 Excel/WPS 中选中的单元格范围和内容，返回地址、数据和工作表名。用于了解用户操作上下文，无需猜测用户要操作的区域",
  parameters: {
    type: "object",
    properties: {},
    required: [],
  },
  riskLevel: "safe",
  requiresApproval: false,
};

export const RANGE_TOOL_DEFINITIONS: ToolDefinition[] = [
  RANGE_READ_DEF,
  RANGE_WRITE_DEF,
  RANGE_CLEAR_DEF,
  SELECTION_GET_DEF,
];
