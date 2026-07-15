/**
 * 直连供应商 + 聚合平台客户端
 *
 * 各厂商的思考等级映射实现。每个子类仅覆盖 applyReasoningConfig() 方法。
 *
 * 直连供应商：
 * - OpenAI — reasoning.effort 参数
 * - DeepSeek — thinking.type + reasoning_effort
 * - 智谱 GLM — thinking.type + reasoning_effort
 * - Kimi — thinking.type 开关
 * - 小米 MiMo — 不传 thinking 参数，使用模型默认
 *
 * 聚合平台：
 * - 讯飞星辰 — enable_thinking + thinking_budget
 * - 百度千帆 — OpenAI 兼容（关闭搜索/引用）
 * - 火山引擎 — thinking.type + budget_tokens
 * - 腾讯云 — OpenAI 兼容（无特殊配置）
 * - 京东云 — OpenAI 兼容（无特殊配置）
 * - 阿里云百炼 — enable_thinking + thinking_budget
 *
 * 关联模块：
 * - openaiCompatibleClient.ts — 基类 OpenAICompatibleClient
 * - aiClientTypes.ts — ReasoningMode 类型
 * - aiClient.ts — 工厂函数 createAIClient 根据 provider/format 路由到本文件中的类
 */

import { OpenAICompatibleClient } from "./openaiCompatibleClient";
import { type ReasoningMode } from "./aiClientTypes";

const THINKING_BUDGET_BY_MODE: Record<string, number> = {
  low: 5000,
  medium: 10000,
  high: 20000,
  max: 20000,
};

export function resolveThinkingBudget(mode?: ReasoningMode): number {
  return THINKING_BUDGET_BY_MODE[mode || "high"] || 10000;
}

// ============================================================
// 直连供应商客户端 — 各厂商思考等级映射
// ============================================================

/** OpenAI — reasoning.effort 参数 */
export class OpenAIClient extends OpenAICompatibleClient {
  protected applyReasoningConfig(body: Record<string, unknown>, mode?: ReasoningMode): void {
    // OpenAI: reasoning.effort，max 映射为 xhigh
    body.reasoning = {
      effort: mode === "max" ? "xhigh" : mode,
    };
  }
}

/** DeepSeek — thinking.type + reasoning_effort */
export class DeepSeekClient extends OpenAICompatibleClient {
  protected applyReasoningConfig(body: Record<string, unknown>, mode?: ReasoningMode): void {
    // DeepSeek: thinking.type + reasoning_effort (只有 high/max 两档)
    body.thinking = { type: "enabled" };
    body.reasoning_effort = mode === "max" ? "max" : "high";
  }
}

/** 智谱 GLM — thinking.type + reasoning_effort */
export class ZhipuClient extends OpenAICompatibleClient {
  protected applyReasoningConfig(body: Record<string, unknown>, mode?: ReasoningMode): void {
    // 智谱: thinking.type + reasoning_effort
    // low/medium 映射为 high，max 保持 max
    body.thinking = { type: "enabled" };
    body.reasoning_effort =
      mode === "low" || mode === "medium" ? "high" : mode === "max" ? "max" : mode || "high";
  }
}

/** Kimi — thinking.type 开关 */
export class KimiClient extends OpenAICompatibleClient {
  protected applyReasoningConfig(body: Record<string, unknown>, _mode?: ReasoningMode): void {
    // Kimi: 只有 enabled/disabled，无等级
    body.thinking = { type: "enabled" };
  }
}

/** 小米 MiMo — thinking.type 开关 */
export class XiaomiClient extends OpenAICompatibleClient {
  protected applyReasoningConfig(_body: Record<string, unknown>, _mode?: ReasoningMode): void {
    // MiMo: 不传 thinking 参数，使用模型默认（开启思考）
    // 关闭时由 streamChat 逻辑处理（不调用 applyReasoningConfig）
  }
}

// ============================================================
// 聚合平台客户端 — 开关型思考模式
// ============================================================

/** 讯飞星辰 — enable_thinking + thinking_budget */
export class XunfeiClient extends OpenAICompatibleClient {
  protected applyReasoningConfig(body: Record<string, unknown>, mode?: ReasoningMode): void {
    body.enable_thinking = true;
    // 聚合平台按档位映射 budget
    body.thinking_budget = resolveThinkingBudget(mode);
  }
}

/** 百度千帆 — 使用 OpenAI 兼容接口 */
export class BaiduClient extends OpenAICompatibleClient {
  protected applyReasoningConfig(body: Record<string, unknown>, _mode?: ReasoningMode): void {
    body.enable_search = false;
    body.enable_citation = false;
  }
}

/** 火山引擎 — thinking.type + budget_tokens */
export class VolcengineClient extends OpenAICompatibleClient {
  protected applyReasoningConfig(body: Record<string, unknown>, mode?: ReasoningMode): void {
    body.thinking = {
      type: "enabled",
      budget_tokens: resolveThinkingBudget(mode),
    };
  }
}

/** 腾讯云 — 使用 OpenAI 兼容接口 */
export class TencentClient extends OpenAICompatibleClient {}

/** 京东云 — 使用 OpenAI 兼容接口 */
export class JDCloudClient extends OpenAICompatibleClient {}

/** 阿里云百炼 — enable_thinking + thinking_budget */
export class AliyunClient extends OpenAICompatibleClient {
  protected applyReasoningConfig(body: Record<string, unknown>, mode?: ReasoningMode): void {
    body.enable_thinking = true;
    body.thinking_budget = resolveThinkingBudget(mode);
  }
}

/** 千问 (Qwen) — 通过 DashScope 兼容模式接入 */
export class QwenClient extends OpenAICompatibleClient {
  protected applyReasoningConfig(body: Record<string, unknown>, mode?: ReasoningMode): void {
    // Qwen 部分模型（如 qwen3-*）支持思考模式
    // 与阿里云百炼共享同一套 API，使用 enable_thinking + thinking_budget
    if (mode && mode !== "off") {
      body.enable_thinking = true;
      body.thinking_budget = resolveThinkingBudget(mode);
    }
  }
}

/** MiniMax — OpenAI 兼容 */
export class MiniMaxClient extends OpenAICompatibleClient {
  protected applyReasoningConfig(body: Record<string, unknown>, mode?: ReasoningMode): void {
    // MiniMax 使用 reasoning_effort 参数（与 OpenAI 类似）
    if (mode && mode !== "off") {
      body.reasoning_config = {
        effort: mode === "max" ? "xhigh" : mode,
      };
    }
  }
}
