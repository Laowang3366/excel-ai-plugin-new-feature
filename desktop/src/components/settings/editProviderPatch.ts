import type { AiProviderConfig, ModelConfig, ReasoningMode } from "../../electronApi";

export interface EditProviderDraft {
  name: string;
  apiFormat: string;
  baseUrl: string;
  apiKey: string;
  model: string;
  contextWindowSize?: number;
  effectiveReasoningMode: ReasoningMode;
  modelConfigs: ModelConfig[];
}

function stripLegacyReasoningOptions(modelConfigs: ModelConfig[]): ModelConfig[] {
  return modelConfigs.map((modelConfig) => {
    const { reasoningOptions: _legacyReasoningOptions, ...rest } = modelConfig;
    return rest;
  });
}

export function buildEditProviderPatch(
  provider: AiProviderConfig,
  draft: EditProviderDraft
): Partial<AiProviderConfig> {
  const patch: Partial<AiProviderConfig> = {};
  const normalizedModelConfigs = stripLegacyReasoningOptions(draft.modelConfigs);

  if (draft.name !== provider.name) patch.name = draft.name;
  if (draft.apiFormat !== (provider.apiFormat || "openai")) patch.apiFormat = draft.apiFormat;
  if (draft.baseUrl !== provider.baseUrl) patch.baseUrl = draft.baseUrl;
  if (draft.apiKey !== provider.apiKey) patch.apiKey = draft.apiKey;
  if (draft.model !== provider.model) patch.model = draft.model;
  if (draft.contextWindowSize !== provider.contextWindowSize) patch.contextWindowSize = draft.contextWindowSize;
  if (draft.effectiveReasoningMode !== (provider.reasoningMode || "off")) {
    patch.reasoningMode = draft.effectiveReasoningMode;
  }
  if (JSON.stringify(normalizedModelConfigs) !== JSON.stringify(provider.modelConfigs || [])) {
    patch.modelConfigs = normalizedModelConfigs;
  }

  return patch;
}
