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

export function buildEditProviderPatch(
  provider: AiProviderConfig,
  draft: EditProviderDraft,
): Partial<AiProviderConfig> {
  const patch: Partial<AiProviderConfig> = {};

  if (draft.name !== provider.name) patch.name = draft.name;
  if (draft.apiFormat !== (provider.apiFormat || "openai")) patch.apiFormat = draft.apiFormat;
  if (draft.baseUrl !== provider.baseUrl) patch.baseUrl = draft.baseUrl;
  if (draft.apiKey !== provider.apiKey) patch.apiKey = draft.apiKey;
  if (draft.model !== provider.model) patch.model = draft.model;
  if (draft.contextWindowSize !== provider.contextWindowSize)
    patch.contextWindowSize = draft.contextWindowSize;
  if (draft.effectiveReasoningMode !== (provider.reasoningMode || "off")) {
    patch.reasoningMode = draft.effectiveReasoningMode;
  }
  if (JSON.stringify(draft.modelConfigs) !== JSON.stringify(provider.modelConfigs || [])) {
    patch.modelConfigs = draft.modelConfigs;
  }

  return patch;
}
