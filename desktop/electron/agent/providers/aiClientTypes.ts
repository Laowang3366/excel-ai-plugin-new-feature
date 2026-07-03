/**
 * AI 客户端类型定义
 *
 * 本文件定义所有 AI 提供商共享的接口和类型。
 * 纯类型文件，无运行时代码，被所有 AI 客户端模块引用。
 *
 * 消费方：
 * - openaiCompatibleClient.ts — 基类使用 ChatMessage、StreamChatParams 等
 * - anthropicClient.ts — 继承基类，使用相同类型
 * - providerClients.ts — 各厂商子类，使用 ReasoningMode
 * - messageBuilder.ts — TurnItem → ChatMessage 转换
 * - agentLoop/ — Agent 主循环使用 StreamChatParams、AIClientConfig
 * - agentLoop/streamCollector.ts — 使用 AIStreamEvent、ChatMessage
 */

import type { ToolDefinition, TokenUsage } from "../shared/types";

// ============================================================
// 多模态内容部分（OpenAI Vision 格式）
// ============================================================

/** 文本内容部分 */
export interface TextContentPart {
  type: "text";
  text: string;
}

/** 图片内容部分（base64 data URI 或 HTTPS URL） */
export interface ImageUrlContentPart {
  type: "image_url";
  image_url: {
    url: string;  // data:image/png;base64,... 或 https://...
    detail?: "low" | "high" | "auto";
  };
}

/** 文件内容部分（PDF 等，base64 data URI） */
export interface FileContentPart {
  type: "file";
  file: {
    filename?: string;
    file_data: string;
    mime_type?: string;
  };
}

/** 聊天消息中的内容部分联合类型 */
export type ContentPart = TextContentPart | ImageUrlContentPart | FileContentPart;

// ============================================================
// 聊天消息格式
// ============================================================

/**
 * 发送给 AI 的聊天消息格式
 * 兼容 OpenAI 和 Anthropic API 格式
 */
export interface ChatMessage {
  role: "system" | "user" | "assistant" | "tool";
  content: string | ContentPart[];
  /** 工具调用的 ID（role=tool 时使用） */
  toolCallId?: string;
  /** 工具调用信息（assistant 消息中的） */
  toolCalls?: ToolCallInfo[];
}

/** AI 返回的工具调用信息 */
export interface ToolCallInfo {
  id: string;
  type: "function";
  function: {
    name: string;
    arguments: string;
  };
}

// ============================================================
// 流式事件
// ============================================================

/** AI 响应的流式事件联合类型 */
export type AIStreamEvent =
  | { type: "text_delta"; delta: string }
  | { type: "reasoning_delta"; delta: string }
  | { type: "reasoning_summary_delta"; delta: string }
  | { type: "tool_call_begin"; toolCallId: string; toolName: string }
  | { type: "tool_call_delta"; toolCallId: string; delta: string }
  | { type: "tool_call_end"; toolCallId: string; toolName: string; arguments: string }
  | { type: "usage"; usage: TokenUsage }
  | { type: "done"; finishReason: string }
  | { type: "error"; error: string };

// ============================================================
// 配置与参数
// ============================================================

/** 统一思考等级枚举 */
export type ReasoningMode = "off" | "low" | "medium" | "high" | "max";

/** 流式聊天请求参数 */
export interface StreamChatParams {
  messages: ChatMessage[];
  tools?: ToolDefinition[];
  systemPrompt?: string;
  maxTokens?: number;
  temperature?: number;
  /** 是否启用思考/推理模式（保留向后兼容，新代码使用 reasoningMode） */
  enableReasoning?: boolean;
  /** 推理力度（保留向后兼容，新代码使用 reasoningMode） */
  reasoningEffort?: "low" | "medium" | "high";
  /** 思考等级，替代 enableReasoning + reasoningEffort */
  reasoningMode?: ReasoningMode;
  /** 信号，用于取消请求 */
  signal?: AbortSignal;
}

/** AI 客户端配置 */
export interface AIClientConfig {
  /** 提供商标识（如 "openai", "anthropic", "deepseek"） */
  provider: string;
  apiKey: string;
  baseUrl: string;
  model: string;
  /** API 协议格式，优先用于客户端路由（参考 Codex model_provider_id） */
  apiFormat?: string;
  /** 自定义请求头 */
  customHeaders?: Record<string, string>;
  /** 是否启用推理模式（保留向后兼容，新代码使用 reasoningMode） */
  enableReasoning?: boolean;
  /** 上下文窗口大小（tokens），用户自定义，未设置时使用默认值 128k */
  contextWindowSize?: number;
  /** 压缩兼容性标识；不同值的模型切换前会先压缩上下文 */
  compHash?: string;
  /** 思考等级，替代 enableReasoning */
  reasoningMode?: ReasoningMode;
  /**
   * 最大输出 tokens。
   * 未设置时根据 reasoningMode 自动推算：
   *   - reasoningMode 为 "off" 或 undefined → 4096
   *   - reasoningMode 为 "low"/"medium" → 8192
   *   - reasoningMode 为 "high"/"max"  → 16384
   */
  maxTokens?: number;
}
