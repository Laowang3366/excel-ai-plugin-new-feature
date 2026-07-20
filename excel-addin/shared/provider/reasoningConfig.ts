import type { ReasoningMode } from "./types";

const REASONING_MODES = new Set<ReasoningMode>([
  "off",
  "low",
  "medium",
  "high",
  "max",
]);

const THINKING_BUDGET_BY_MODE: Record<string, number> = {
  low: 5000,
  medium: 10000,
  high: 20000,
  max: 20000,
};

export function isReasoningMode(value: unknown): value is ReasoningMode {
  return typeof value === "string" && REASONING_MODES.has(value as ReasoningMode);
}

/** Desktop-compatible thinking budget for aggregation / Qwen-style APIs. */
export function resolveThinkingBudget(mode?: ReasoningMode): number {
  return THINKING_BUDGET_BY_MODE[mode || "high"] || 10000;
}

/**
 * OpenAI Responses: off must send effort=none explicitly (do not omit).
 * max → xhigh; invalid/undefined → omit entirely.
 */
export function applyResponsesReasoningConfig(
  body: Record<string, unknown>,
  mode?: ReasoningMode | string,
): void {
  if (!isReasoningMode(mode)) return;
  const effort = mode === "off" ? "none" : mode === "max" ? "xhigh" : mode;
  body.reasoning = {
    effort,
    ...(effort !== "none" ? { summary: "auto" } : {}),
  };
}

/**
 * Anthropic Messages: adaptive thinking + output_config.effort.
 * off / invalid → omit.
 */
export function applyAnthropicReasoningConfig(
  body: Record<string, unknown>,
  mode?: ReasoningMode | string,
): void {
  if (!isReasoningMode(mode) || mode === "off") return;
  body.thinking = { type: "adaptive" };
  body.output_config = {
    effort: mode === "max" ? "xhigh" : mode,
  };
}

/**
 * Chat Completions vendor mapping mirrors desktop providerClients.ts.
 * Only applied when mode is a known non-off value.
 * Unknown / custom providers get no OpenAI-only fields.
 */
export function applyChatCompletionsReasoningConfig(
  body: Record<string, unknown>,
  provider: string | undefined,
  mode?: ReasoningMode | string,
): void {
  if (!isReasoningMode(mode) || mode === "off") return;

  const vendor = typeof provider === "string" ? provider.trim().toLowerCase() : "";

  switch (vendor) {
    case "openai":
      // OpenAI: reasoning.effort; max → xhigh
      body.reasoning = {
        effort: mode === "max" ? "xhigh" : mode,
      };
      break;

    case "deepseek":
      // DeepSeek: thinking.type + reasoning_effort (high/max only)
      body.thinking = { type: "enabled" };
      body.reasoning_effort = mode === "max" ? "max" : "high";
      break;

    case "zhipu":
      // 智谱: thinking.type + reasoning_effort; low/medium → high
      body.thinking = { type: "enabled" };
      body.reasoning_effort =
        mode === "low" || mode === "medium"
          ? "high"
          : mode === "max"
            ? "max"
            : mode || "high";
      break;

    case "kimi":
      // Kimi: enabled/disabled only
      body.thinking = { type: "enabled" };
      break;

    case "xiaomi":
      // MiMo: leave defaults (no thinking fields)
      break;

    case "xunfei":
      body.enable_thinking = true;
      body.thinking_budget = resolveThinkingBudget(mode);
      break;

    case "baidu":
      body.enable_search = false;
      body.enable_citation = false;
      break;

    case "volcengine":
      body.thinking = {
        type: "enabled",
        budget_tokens: resolveThinkingBudget(mode),
      };
      break;

    case "tencent":
    case "jdcloud":
    case "custom":
      // OpenAI-compatible with no special reasoning fields
      break;

    case "aliyun":
      body.enable_thinking = true;
      body.thinking_budget = resolveThinkingBudget(mode);
      break;

    case "qwen":
      body.enable_thinking = true;
      body.thinking_budget = resolveThinkingBudget(mode);
      break;

    case "minimax":
      body.reasoning_config = {
        effort: mode === "max" ? "xhigh" : mode,
      };
      break;

    default:
      // Unknown vendor: do not send OpenAI-only fields
      break;
  }
}
