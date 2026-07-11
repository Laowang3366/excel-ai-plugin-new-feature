import type { AiProviderConfig, ModelConfig, ReasoningMode } from "../../electronApi";
import type { ProviderTemplate } from "../../store/settingsProviderTemplates";

export interface AddProviderDraft {
  selectedTemplateId: string;
  name: string;
  apiFormat: string;
  baseUrl: string;
  apiKey?: string;
  model: string;
  contextWindowSize?: number;
  reasoningMode: ReasoningMode;
  modelConfigs: ModelConfig[];
}

export interface BuildProviderConfigInput {
  id: string;
  draft: AddProviderDraft;
  selectedTemplate: ProviderTemplate | null;
  effectiveReasoningMode: ReasoningMode;
  customProviderName: string;
  resolvedModel?: string;
}

export function createEmptyProviderDraft(): AddProviderDraft {
  return {
    selectedTemplateId: "",
    name: "",
    apiFormat: "openai",
    baseUrl: "",
    model: "",
    contextWindowSize: undefined,
    reasoningMode: "off",
    modelConfigs: [],
  };
}

export function providerDraftFromTemplate(template: ProviderTemplate): AddProviderDraft {
  return {
    selectedTemplateId: template.id,
    name: template.name,
    apiFormat: template.apiFormat,
    baseUrl: template.baseUrl,
    model: template.defaultModel,
    contextWindowSize: template.defaultContextWindowSize,
    reasoningMode: template.defaultReasoningMode || "off",
    modelConfigs: [],
  };
}

export function buildProviderConfigFromDraft({
  id,
  draft,
  selectedTemplate,
  effectiveReasoningMode,
  customProviderName,
  resolvedModel,
}: BuildProviderConfigInput): AiProviderConfig {
  const finalModel = resolvedModel || draft.model || selectedTemplate?.defaultModel || "";

  return {
    id,
    name: draft.name || selectedTemplate?.name || customProviderName,
    provider: selectedTemplate?.provider || "custom",
    apiKey: draft.apiKey || "",
    baseUrl: draft.baseUrl,
    model: finalModel,
    defaultBaseUrl: selectedTemplate?.baseUrl || draft.baseUrl,
    defaultModel: selectedTemplate?.defaultModel || "",
    reasoningMode: effectiveReasoningMode,
    apiFormat: draft.apiFormat,
    models: selectedTemplate?.presetModels || undefined,
    modelConfigs: draft.modelConfigs.length > 0 ? draft.modelConfigs : undefined,
    contextWindowSize:
      draft.contextWindowSize && draft.contextWindowSize > 0 ? draft.contextWindowSize : undefined,
  };
}
