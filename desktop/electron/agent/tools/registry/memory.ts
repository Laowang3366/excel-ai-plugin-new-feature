import type { ToolDefinition } from "../../shared/types";
import { TOOL_WRITABLE_MEMORY_KINDS } from "../../memory/longTerm/memoryTypes";

export const MEMORY_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "memory.write",
    description:
      "写入用户长期记忆。仅用于用户明确偏好、长期约束、纠正、文档风格偏好、操作方式偏好和低敏文件印象；不要写入文件正文、表格明细、临时路径或内部工具统计。",
    parameters: {
      type: "object",
      properties: {
        kind: {
          type: "string",
          enum: [...TOOL_WRITABLE_MEMORY_KINDS],
        },
        namespace: { type: "string", default: "global" },
        content: { type: "string" },
        summary: { type: "string" },
        confidence: { type: "number" },
      },
      required: ["kind", "content"],
    },
    riskLevel: "safe",
    requiresApproval: false,
  },
  {
    name: "memory.search",
    description:
      "搜索用户可见长期记忆，用于了解用户偏好、长期约束、纠正、文档风格偏好、操作方式偏好和过往文件印象。",
    parameters: {
      type: "object",
      properties: {
        query: { type: "string" },
        namespace: { type: "string" },
        kind: { type: "string" },
        limit: { type: "number", default: 10 },
      },
      required: ["query"],
    },
    riskLevel: "safe",
    requiresApproval: false,
  },
  {
    name: "memory.list",
    description: "列出用户可见长期记忆摘要。",
    parameters: {
      type: "object",
      properties: {
        namespace: { type: "string" },
      },
    },
    riskLevel: "safe",
    requiresApproval: false,
  },
  {
    name: "memory.delete",
    description:
      "按 memoryId 删除/停用一条用户可见长期记忆。需要先通过 memory.list 或 memory.search 确认要删除的记忆 ID；不会删除知识库内容。",
    parameters: {
      type: "object",
      properties: {
        memoryId: { type: "string" },
      },
      required: ["memoryId"],
    },
    riskLevel: "safe",
    requiresApproval: false,
  },
];
