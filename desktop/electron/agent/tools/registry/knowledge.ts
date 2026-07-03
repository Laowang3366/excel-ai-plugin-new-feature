/**
 * 知识库工具定义
 *
 * 包含本地知识库搜索工具。
 */

import type { ToolDefinition } from "../../shared/types";

/** 知识库搜索 */
const KNOWLEDGE_SEARCH_DEF: ToolDefinition = {
  name: "knowledge.search",
  description: "搜索本地的知识库，获取与当前任务相关的历史信息，包括工作簿结构、字段含义、公式规则、操作记录等。当需要了解表格结构、字段定义、过往操作时可使用此工具",
  parameters: {
    type: "object",
    properties: {
      query: { type: "string", description: "搜索关键词，建议包含表名、字段名或操作类型，如「销售表 金额列含义」" },
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
