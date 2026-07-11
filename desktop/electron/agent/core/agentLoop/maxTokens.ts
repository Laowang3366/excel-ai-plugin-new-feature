import type { AIClientConfig } from "../../providers/aiClient";

/**
 * Agent 输出 token 预算计算。
 *
 * 关联模块：
 * - agentLoop.ts: 构建模型流式请求时使用。
 * - providers/aiClient: 提供上下文窗口和推理模式配置。
 */
export function resolveMaxTokens(aiConfig: AIClientConfig): number {
  if (aiConfig.maxTokens !== undefined && aiConfig.maxTokens > 0) {
    return aiConfig.maxTokens;
  }

  const ctxSize = aiConfig.contextWindowSize || 128_000;
  const baseFromCtx = Math.min(Math.max(Math.floor(ctxSize * 0.06), 4_096), 16_384);
  const mode = aiConfig.reasoningMode || "off";

  switch (mode) {
    case "max":
      return Math.min(Math.max(baseFromCtx * 2, 32_768), ctxSize * 0.25);
    case "high":
      return Math.min(Math.max(baseFromCtx, 16_384), ctxSize * 0.15);
    case "medium":
    case "low":
      return Math.min(Math.max(baseFromCtx, 8_192), ctxSize * 0.1);
    case "off":
    default:
      return Math.min(Math.max(baseFromCtx, 4_096), ctxSize * 0.08);
  }
}
