/**
 * 公式工具定义
 *
 * 包含公式上下文读取工具。
 */

import type { ToolDefinition } from "../../shared/types";

/** 获取公式上下文 */
const FORMULA_CONTEXT_DEF: ToolDefinition = {
  name: "formula.context",
  description:
    "获取工作表指定范围中的公式信息，返回每个含公式单元格的地址、公式和值。用于理解已有公式逻辑、检查公式引用关系、辅助生成新公式",
  parameters: {
    type: "object",
    properties: {
      sheetName: { type: "string", description: "工作表名称" },
      range: { type: "string", description: "范围（可选，默认扫描已用区域）" },
    },
    required: ["sheetName"],
  },
  riskLevel: "safe",
  requiresApproval: false,
};

export const FORMULA_TOOL_DEFINITIONS: ToolDefinition[] = [FORMULA_CONTEXT_DEF];
