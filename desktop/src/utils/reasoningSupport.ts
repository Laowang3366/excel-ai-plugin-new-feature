import type { AiProviderConfig, ModelConfig, ReasoningMode } from "../electronApi";

export interface ReasoningTemplateLike {
  provider?: string;
  apiFormat?: string;
  defaultReasoningMode?: ReasoningMode;
  reasoningOptions?: Array<{ value: ReasoningMode | string; label: string }>;
}

export const REASONING_FULL: ReasoningMode[] = ["off", "low", "medium", "high", "max"];
export const REASONING_HIGH_MAX: ReasoningMode[] = ["off", "high", "max"];
export const REASONING_TOGGLE: ReasoningMode[] = ["off", "high"];

const LABELS_ZH: Record<ReasoningMode, string> = {
  off: "关闭",
  low: "低",
  medium: "中",
  high: "高",
  max: "极高",
};

const LABELS_EN: Record<ReasoningMode, string> = {
  off: "Off",
  low: "Low",
  medium: "Medium",
  high: "High",
  max: "Max",
};

function toReasoningMode(value: unknown): ReasoningMode | null {
  return value === "off" ||
    value === "low" ||
    value === "medium" ||
    value === "high" ||
    value === "max"
    ? value
    : null;
}

function uniqueModes(values: ReasoningMode[]): ReasoningMode[] {
  return values.filter((value, index) => values.indexOf(value) === index);
}

function templateModes(template?: ReasoningTemplateLike | null): ReasoningMode[] {
  const values = (template?.reasoningOptions || [])
    .map((option) => toReasoningMode(option.value))
    .filter((value): value is ReasoningMode => Boolean(value));
  return values.length >= 2 ? uniqueModes(values) : REASONING_TOGGLE;
}

function modelLooksLike(pattern: RegExp, model: string, provider: string): boolean {
  return pattern.test(`${provider} ${model}`.toLowerCase());
}

export function resolveReasoningOptionValues(
  provider: Pick<AiProviderConfig, "provider" | "apiFormat" | "model">,
  template?: ReasoningTemplateLike | null,
  modelConfig?: Pick<ModelConfig, "name"> | null,
): ReasoningMode[] {
  const providerId = (provider.provider || template?.provider || "").toLowerCase();
  const apiFormat = (provider.apiFormat || template?.apiFormat || "").toLowerCase();
  const modelName = (modelConfig?.name || provider.model || "").toLowerCase();

  if (
    apiFormat === "responses" ||
    providerId === "openai" ||
    modelLooksLike(/(^|\s)(gpt-|o\d|codex)/, modelName, providerId)
  ) {
    return REASONING_FULL;
  }

  if (providerId === "anthropic" || modelLooksLike(/claude/, modelName, providerId)) {
    return REASONING_FULL;
  }

  if (modelLooksLike(/deepseek|glm|zhipu|qwen|minimax/, modelName, providerId)) {
    return REASONING_HIGH_MAX;
  }

  if (modelLooksLike(/kimi|moonshot|mimo|xiaomi/, modelName, providerId)) {
    return REASONING_TOGGLE;
  }

  return templateModes(template);
}

export function defaultReasoningModeForOptions(
  options: ReasoningMode[],
  preferred?: ReasoningMode,
): ReasoningMode {
  if (preferred && options.includes(preferred)) return preferred;
  if (options.includes("medium")) return "medium";
  if (options.includes("high")) return "high";
  return "off";
}

export function coerceReasoningMode(
  mode: ReasoningMode | undefined,
  options: ReasoningMode[],
  preferredDefault?: ReasoningMode,
): ReasoningMode {
  if (mode && options.includes(mode)) return mode;
  return defaultReasoningModeForOptions(options, preferredDefault);
}

export function buildReasoningOptions(
  options: ReasoningMode[],
  language: "zh-CN" | "en-US",
): Array<{ value: ReasoningMode; label: string }> {
  const labels = language === "zh-CN" ? LABELS_ZH : LABELS_EN;
  return options.map((value) => ({ value, label: labels[value] }));
}

export function formatReasoningOptionLabels(
  options: ReasoningMode[],
  language: "zh-CN" | "en-US",
): string {
  const labels = language === "zh-CN" ? LABELS_ZH : LABELS_EN;
  return options.map((value) => labels[value]).join(" / ");
}

export function normalizeProviderReasoningConfig(
  provider: AiProviderConfig,
  template?: ReasoningTemplateLike | null,
): AiProviderConfig {
  const modelConfigs = provider.modelConfigs?.map((modelConfig) => {
    const options = resolveReasoningOptionValues(provider, template, modelConfig);
    const reasoningMode = modelConfig.reasoningMode && options.includes(modelConfig.reasoningMode)
      ? modelConfig.reasoningMode
      : undefined;
    const { reasoningOptions: _legacyReasoningOptions, ...rest } = modelConfig;
    return reasoningMode ? { ...rest, reasoningMode } : rest;
  });

  const activeModelConfig = modelConfigs?.find((modelConfig) => modelConfig.name === provider.model);
  const options = resolveReasoningOptionValues(provider, template, activeModelConfig);
  const reasoningMode = coerceReasoningMode(
    activeModelConfig?.reasoningMode || provider.reasoningMode,
    options,
    template?.defaultReasoningMode,
  );

  return {
    ...provider,
    reasoningMode,
    enableReasoning: reasoningMode !== "off" ? true : undefined,
    modelConfigs,
  };
}
