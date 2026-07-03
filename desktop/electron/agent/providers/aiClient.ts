/**
 * AI 客户端层 — barrel re-export
 *
 * 支持 OpenAI/Anthropic/国内厂商的统一接口。
 * 所有类型、类和工厂函数从子模块统一导出，保持向后兼容。
 *
 * 子模块：
 * - aiClientTypes.ts    — 类型定义（ChatMessage、StreamChatParams 等）
 * - openaiCompatibleClient.ts — OpenAI 兼容基类 + SSE 解析
 * - anthropicClient.ts  — Anthropic 客户端（消息格式/thinking blocks）
 * - providerClients.ts  — 11 个厂商子类（DeepSeek/Kimi/智谱等）
 * - aiClientFactory.ts  — createAIClient 工厂函数
 * - messageBuilder.ts   — TurnItem → ChatMessage 转换器
 */

// ── 类型 ──
export {
  type TextContentPart,
  type ImageUrlContentPart,
  type FileContentPart,
  type ContentPart,
  type ChatMessage,
  type ToolCallInfo,
  type AIStreamEvent,
  type ReasoningMode,
  type StreamChatParams,
  type AIClientConfig,
} from "./aiClientTypes";

// ── TokenUsage 从 types 透传（streamCollector 等模块依赖此路径） ──
export type { TokenUsage } from "../shared/types";

// ── 工具名称清洗 ──
export { sanitizeToolName, desanitizeToolName } from "./openaiCompatibleClient";

// ── 基类 ──
export { OpenAICompatibleClient } from "./openaiCompatibleClient";
export { OpenAIResponsesClient } from "./openaiResponsesClient";

// ── Anthropic 客户端 ──
export { AnthropicClient } from "./anthropicClient";

// ── 厂商子类 ──
export {
  OpenAIClient,
  DeepSeekClient,
  ZhipuClient,
  KimiClient,
  XiaomiClient,
  XunfeiClient,
  BaiduClient,
  VolcengineClient,
  TencentClient,
  JDCloudClient,
  AliyunClient,
} from "./providerClients";

// ── 工厂函数 ──
export { createAIClient } from "./aiClientFactory";

// ── 消息构建 ──
export { turnItemsToChatMessages } from "../shared/messageBuilder";
