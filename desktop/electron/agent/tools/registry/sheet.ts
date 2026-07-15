/**
 * 工作表工具定义
 *
 * 包含工作表增删改移动等结构操作。
 */

import type { ToolDefinition } from "../../shared/types";

/** 工作表操作 */
const SHEET_OPERATION_DEF: ToolDefinition = {
  name: "sheet.operation",
  description:
    "对工作表执行增删改操作：新建(add)、重命名(rename)、删除(delete)、复制(copy)、移动(move)",
  parameters: {
    type: "object",
    properties: {
      operation: {
        type: "string",
        enum: ["rename", "move", "delete", "copy", "add"],
        description: "操作类型",
      },
      sheetName: { type: "string", description: "目标工作表名称" },
      newName: { type: "string", description: "新名称（rename 时使用）" },
      position: { type: "integer", minimum: 1, description: "位置（move 时使用）" },
    },
    required: ["operation", "sheetName"],
  },
  riskLevel: "moderate",
  requiresApproval: true,
  isFileDeletion: true,
};

export const SHEET_TOOL_DEFINITIONS: ToolDefinition[] = [SHEET_OPERATION_DEF];
