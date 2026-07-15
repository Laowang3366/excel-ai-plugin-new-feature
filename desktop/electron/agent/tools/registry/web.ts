import type { ToolDefinition } from "../../shared/types";

const WEB_SEARCH_DEF: ToolDefinition = {
  name: "web.search",
  description:
    "联网搜索公开网页信息，用于回答需要实时信息、最新资料、外部网页事实或来源链接的问题。搜索结果只用于回答；只有用户明确要求保存时才写入知识库",
  parameters: {
    type: "object",
    properties: {
      query: { type: "string", description: "搜索关键词或问题" },
      maxResults: {
        type: "integer",
        minimum: 1,
        maximum: 10,
        description: "返回结果数量，默认5，最多10",
        default: 5,
      },
      freshness: {
        type: "string",
        enum: ["day", "week", "month", "year", "any"],
        description: "可选时间范围偏好；不同搜索后端可能部分支持",
        default: "any",
      },
    },
    required: ["query"],
  },
  riskLevel: "moderate",
  requiresApproval: true,
  isDataEgress: true,
};

export const WEB_TOOL_DEFINITIONS: ToolDefinition[] = [WEB_SEARCH_DEF];
