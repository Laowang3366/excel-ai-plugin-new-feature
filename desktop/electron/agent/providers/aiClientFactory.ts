/**
 * AI 客户端工厂
 *
 * 参考 Codex 的 model_provider_id 概念：
 * 优先使用 apiFormat 路由，fallback 到 provider 字段。
 * 对于 apiFormat="openai" 的直连供应商，按 provider 进一步路由到专用客户端。
 *
 * 关联模块：
 * - aiClientTypes.ts — AIClientConfig 配置类型
 * - openaiCompatibleClient.ts — 基类 OpenAICompatibleClient
 * - anthropicClient.ts — Anthropic 客户端
 * - providerClients.ts — 各厂商子类（OpenAI/DeepSeek/智谱/Kimi 等）
 * - agentLoop/agentLoop.ts — 调用本工厂创建 AI 客户端
 */

import { type AIClientConfig } from "./aiClientTypes";
import { OpenAICompatibleClient } from "./openaiCompatibleClient";
import { OpenAIResponsesClient } from "./openaiResponsesClient";
import { AnthropicClient } from "./anthropicClient";
import {
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
  QwenClient,
  MiniMaxClient,
} from "./providerClients";

/**
 * 创建 AI 客户端实例
 *
 * 路由规则：
 * 1. 优先使用 apiFormat 字段确定协议格式（openai / anthropic / responses）
 * 2. 对于 apiFormat="openai"，按 provider 字段进一步路由到专用子类
 * 3. 对于 apiFormat="responses"，使用 OpenAI Responses 格式
 * 4. 未知格式回退到 OpenAICompatibleClient
 */
export function createAIClient(config: AIClientConfig): OpenAICompatibleClient {
  const format = config.apiFormat || config.provider;

  switch (format) {
    case "anthropic":
      return new AnthropicClient(config);

    case "responses":
      return new OpenAIResponsesClient(config);

    case "openai":
    default:
      // 按 provider 路由到专用子类（各家在 OpenAI 兼容协议上的扩展差异）
      switch (config.provider) {
        case "openai":
          return new OpenAIClient(config);
        case "deepseek":
          return new DeepSeekClient(config);
        case "zhipu":
          return new ZhipuClient(config);
        case "kimi":
          return new KimiClient(config);
        case "xiaomi":
          return new XiaomiClient(config);
        case "xunfei":
          return new XunfeiClient(config);
        case "baidu":
          return new BaiduClient(config);
        case "volcengine":
          return new VolcengineClient(config);
        case "tencent":
          return new TencentClient(config);
        case "jdcloud":
          return new JDCloudClient(config);
        case "aliyun":
          return new AliyunClient(config);
        case "qwen":
          return new QwenClient(config);
        case "minimax":
          return new MiniMaxClient(config);
        default:
          return new OpenAICompatibleClient(config);
      }
  }
}
