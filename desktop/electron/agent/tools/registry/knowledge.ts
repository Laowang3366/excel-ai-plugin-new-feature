/**
 * 知识库工具定义
 *
 * 包含本地知识库搜索工具。
 */

import type { ToolDefinition } from "../../shared/types";

/** 知识库搜索 */
const KNOWLEDGE_SEARCH_DEF: ToolDefinition = {
  name: "knowledge.search",
  description: "搜索本地知识库，获取与当前任务相关的历史项目知识、字段口径、公式规则、模板规范和操作记录。先读取当前文件/数据并判断场景难度；简单问答、单步操作、直观格式调整无需调用，只有中高复杂度或明确依赖业务资料/历史规则时再用场景摘要检索",
  parameters: {
    type: "object",
    properties: {
      query: { type: "string", description: "场景摘要式搜索词，建议包含任务类型、文件/表/页/章节、字段名、样例值、业务口径和目标输出，如「销售汇总公式 订单表 区域 金额 月份 多条件汇总」" },
      topK: { type: "number", description: "返回结果数量，默认5", default: 5 },
    },
    required: ["query"],
  },
  riskLevel: "safe",
  requiresApproval: false,
};

/** 知识库写入 */
const KNOWLEDGE_WRITE_DEF: ToolDefinition = {
  name: "knowledge.write",
  description: "把用户明确要求保存为知识库的内容写入本地知识库，供后续 knowledge.search 检索。仅在用户说“写入知识库”“记到知识库”“保存为知识”等明确触发时使用，不要替代 memory.write",
  parameters: {
    type: "object",
    properties: {
      content: { type: "string", description: "要写入知识库的正文，必须是用户明确希望沉淀的知识内容" },
      title: { type: "string", description: "知识条目标题，可选；未提供时从正文首行生成" },
      tags: {
        type: "array",
        description: "可选标签，例如项目名、业务域、文件名",
        items: { type: "string" },
      },
      sourceName: { type: "string", description: "可选来源文件名；默认自动生成 note-*.md" },
    },
    required: ["content"],
  },
  riskLevel: "safe",
  requiresApproval: false,
};

export const KNOWLEDGE_TOOL_DEFINITIONS: ToolDefinition[] = [
  KNOWLEDGE_SEARCH_DEF,
  KNOWLEDGE_WRITE_DEF,
];
