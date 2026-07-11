/**
 * 公式工具定义
 *
 * 包含公式上下文读取、解题契约、验收和函数搜索工具。
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

const FORMULA_PREPARE_DEF: ToolDefinition = {
  name: "formula.prepare",
  description: "公式任务读取数据源和指定参考样例后，提交结构化解题与验收契约。必须先判断场景、输入输出形状、业务粒度、业务键和必要变换；若口语需求存在会影响结果的多种解释，使用 needs_clarification 并提出一个明确问题，不得猜测后写入。参考样例为空时必须提供可执行的无样例验收检查。",
  parameters: {
    type: "object",
    properties: {
      status: {
        type: "string",
        enum: ["ready", "needs_clarification"],
        description: "ready=任务契约明确，可以继续；needs_clarification=存在影响结果的实质歧义，必须先询问用户",
      },
      scenario: { type: "string", description: "场景判断，如分组聚合、查找映射、文本提取、数量展开或重塑" },
      inputShape: { type: "string", description: "读取后判断的输入形状" },
      outputShape: { type: "string", description: "根据需求和参考样例判断的输出形状" },
      inputGrain: { type: "string", description: "输入中一行代表的业务粒度" },
      outputGrain: { type: "string", description: "输出中一行或一个单元格代表的业务粒度" },
      businessKeys: { type: "array", items: { type: "string" }, description: "筛选、匹配、分组或排序使用的业务键" },
      transformChain: { type: "array", items: { type: "string" }, description: "从输入到输出的必要变换链，不添加无关阶段" },
      constraints: { type: "array", items: { type: "string" }, description: "Excel/WPS、动态数组、兼容性和目标区域等约束" },
      acceptanceChecks: {
        type: "array",
        description: "写入后需要执行的验收检查。无参考样例时除 no_excel_error 外，至少再提供一项 shape、unique_key、row_count、aggregate_reconciliation、sort_order、lookup_consistency、pattern_match、boundary 或 spot_check",
        items: {
          type: "object",
          properties: {
            type: {
              type: "string",
              enum: ["no_excel_error", "shape", "unique_key", "row_count", "aggregate_reconciliation", "sort_order", "lookup_consistency", "pattern_match", "boundary", "spot_check"],
            },
            description: { type: "string" },
            required: { type: "boolean" },
            params: { type: "object", description: "检查参数，列号使用从 1 开始的数字" },
          },
          required: ["type", "description"],
        },
      },
      assumptions: { type: "array", items: { type: "string" }, description: "不影响核心结果的次要假设；实质歧义不能写在这里" },
      clarificationQuestion: { type: "string", description: "needs_clarification 时向用户提出的问题" },
    },
    required: ["status"],
  },
  riskLevel: "safe",
  requiresApproval: false,
};

const FORMULA_VERIFY_DEF: ToolDefinition = {
  name: "formula.verify",
  description: "对刚写入的公式执行运行时验收，返回实际溢出范围、尺寸、错误值、样例差异和无样例不变量结果。通常由运行时自动调用；不得在 range.write 之前调用。",
  parameters: { type: "object", properties: {}, required: [] },
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
  FORMULA_PREPARE_DEF,
  FORMULA_VERIFY_DEF,
  FORMULA_SEARCH_DEF,
];
