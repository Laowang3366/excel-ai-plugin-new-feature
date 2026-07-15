/**
 * 知识库工具定义
 *
 * 包含本地知识库搜索工具。
 */

import type { ToolDefinition } from "../../shared/types";

/** 知识库搜索 */
const KNOWLEDGE_SEARCH_DEF: ToolDefinition = {
  name: "knowledge.search",
  description: "搜索本地知识库中的项目资料、业务口径、模板规范和历史规则。先读取当前文件或数据，仅在任务依赖已沉淀知识或用户明确要求时调用；返回片段用于补充当前文件事实，不得把最相似案例直接当作答案",
  parameters: {
    type: "object",
    properties: {
      query: { type: "string", description: "结构化搜索词，包含任务、业务对象、关键字段、约束和目标；样例值只作补充" },
      topK: { type: "integer", minimum: 1, maximum: 50, description: "返回结果数量，默认5，最多50", default: 5 },
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

const KNOWLEDGE_LIST_SOURCES_DEF: ToolDefinition = {
  name: "knowledge.listSources",
  description: "列出已经索引到本地知识库的来源文件，用于在修改、追加或删除知识库内容前确认 sourcePath。普通检索仍优先使用 knowledge.search。",
  parameters: {
    type: "object",
    properties: {},
  },
  riskLevel: "safe",
  requiresApproval: false,
};

const KNOWLEDGE_UPDATE_SOURCE_DEF: ToolDefinition = {
  name: "knowledge.updateSource",
  description:
    "修改已经索引的可写文本知识来源。仅在用户明确要求修改或追加知识库内容时使用；调用前应先用 knowledge.listSources 或 knowledge.search 确认 sourcePath。replace 会替换该来源内容，append 会追加到该来源末尾，并重建该来源索引。",
  parameters: {
    type: "object",
    properties: {
      sourcePath: { type: "string", description: "要修改的已索引知识来源绝对路径，通常来自 knowledge.listSources 的返回值。" },
      operation: {
        type: "string",
        enum: ["replace", "append"],
        description: "replace 表示替换该来源内容；append 表示追加到该来源末尾。",
      },
      content: { type: "string", description: "要写入或追加的知识库正文内容。" },
      title: { type: "string", description: "可选标题；replace 写入 Markdown 来源时会作为标题使用。" },
      tags: {
        type: "array",
        description: "可选标签。",
        items: { type: "string" },
      },
    },
    required: ["sourcePath", "operation", "content"],
  },
  riskLevel: "moderate",
  requiresApproval: false,
};

const KNOWLEDGE_DELETE_SOURCE_DEF: ToolDefinition = {
  name: "knowledge.deleteSource",
  description:
    "删除指定来源在知识库中的索引内容，不删除磁盘上的原始文件。仅在用户明确要求从知识库删除某个来源或内容时使用；调用前应先确认 sourcePath。",
  parameters: {
    type: "object",
    properties: {
      sourcePath: { type: "string", description: "要从知识库索引中删除的来源绝对路径，通常来自 knowledge.listSources 的返回值。" },
    },
    required: ["sourcePath"],
  },
  riskLevel: "moderate",
  requiresApproval: false,
};

export const KNOWLEDGE_TOOL_DEFINITIONS: ToolDefinition[] = [
  KNOWLEDGE_SEARCH_DEF,
  KNOWLEDGE_LIST_SOURCES_DEF,
  KNOWLEDGE_WRITE_DEF,
  KNOWLEDGE_UPDATE_SOURCE_DEF,
  KNOWLEDGE_DELETE_SOURCE_DEF,
];
