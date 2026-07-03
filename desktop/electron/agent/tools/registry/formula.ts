/**
 * 公式工具定义
 *
 * 包含公式上下文读取和函数搜索工具。
 */

import type { ToolDefinition } from "../../shared/types";

/** 获取公式上下文 */
const FORMULA_CONTEXT_DEF: ToolDefinition = {
  name: "formula.context",
  description: "获取工作表指定范围中的公式信息，返回每个含公式单元格的地址、公式和值。用于理解已有公式逻辑、检查公式引用关系、辅助生成新公式",
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

/** 公式搜索 */
const FORMULA_SEARCH_DEF: ToolDefinition = {
  name: "formula.search",
  description: "搜索 Excel 内置函数库，返回匹配函数的名称、语法和用法说明。用于查找合适的函数、了解函数参数和示例",
  parameters: {
    type: "object",
    properties: {
      query: { type: "string", description: "搜索关键词（函数名或功能描述）" },
      category: { type: "string", description: "函数类别（math/statistics/lookup/text/logic/date/financial/array）" },
    },
    required: ["query"],
  },
  riskLevel: "safe",
  requiresApproval: false,
};

export const FORMULA_TOOL_DEFINITIONS: ToolDefinition[] = [
  FORMULA_CONTEXT_DEF,
  FORMULA_SEARCH_DEF,
];
