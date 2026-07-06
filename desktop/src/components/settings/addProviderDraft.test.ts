import { describe, expect, it } from "vitest";
import type { ProviderTemplate } from "../../store/settingsProviderTemplates";
import { buildProviderConfigFromDraft, createEmptyProviderDraft, providerDraftFromTemplate } from "./addProviderDraft";

const template: ProviderTemplate = {
  id: "openai",
  name: "OpenAI",
  provider: "openai",
  baseUrl: "https://api.openai.com/v1",
  defaultModel: "gpt-5.4",
  apiFormat: "openai",
  presetModels: ["gpt-5.4"],
  defaultContextWindowSize: 256000,
  category: "direct",
  reasoningOptions: [{ value: "off", label: "Off" }, { value: "medium", label: "Medium" }],
  defaultReasoningMode: "medium",
};

describe("addProviderDraft", () => {
  it("creates empty and template-based drafts", () => {
    expect(createEmptyProviderDraft()).toEqual({
      selectedTemplateId: "",
      name: "",
      apiFormat: "openai",
      baseUrl: "",
      model: "",
      contextWindowSize: undefined,
      reasoningMode: "off",
      modelConfigs: [],
    });

    expect(providerDraftFromTemplate(template)).toEqual({
      selectedTemplateId: "openai",
      name: "OpenAI",
      apiFormat: "openai",
      baseUrl: "https://api.openai.com/v1",
      model: "gpt-5.4",
      contextWindowSize: 256000,
      reasoningMode: "medium",
      modelConfigs: [],
    });
  });

  it("builds provider configs from the current form draft", () => {
    const config = buildProviderConfigFromDraft({
      id: "provider_1",
      draft: {
        ...providerDraftFromTemplate(template),
        model: "gpt-5.4",
        apiKey: "sk-test",
        modelConfigs: [{ name: "gpt-5.4", contextWindowSize: 256000 }],
      },
      selectedTemplate: template,
      effectiveReasoningMode: "medium",
      customProviderName: "Custom Provider",
    });

    expect(config).toMatchObject({
      id: "provider_1",
      name: "OpenAI",
      provider: "openai",
      apiKey: "sk-test",
      baseUrl: "https://api.openai.com/v1",
      model: "gpt-5.4",
      defaultBaseUrl: "https://api.openai.com/v1",
      defaultModel: "gpt-5.4",
      enableReasoning: true,
      reasoningMode: "medium",
      apiFormat: "openai",
      models: ["gpt-5.4"],
      contextWindowSize: 256000,
    });
    expect(config.modelConfigs).toEqual([{ name: "gpt-5.4", contextWindowSize: 256000 }]);
  });

  it("falls back to custom provider defaults without optional empty fields", () => {
    const config = buildProviderConfigFromDraft({
      id: "provider_custom",
      draft: {
        ...createEmptyProviderDraft(),
        apiKey: "sk-custom",
        baseUrl: "https://example.com/v1",
      },
      selectedTemplate: null,
      effectiveReasoningMode: "off",
      customProviderName: "Custom Provider",
    });

    expect(config).toEqual({
      id: "provider_custom",
      name: "Custom Provider",
      provider: "custom",
      apiKey: "sk-custom",
      baseUrl: "https://example.com/v1",
      model: "",
      defaultBaseUrl: "https://example.com/v1",
      defaultModel: "",
      enableReasoning: undefined,
      reasoningMode: "off",
      apiFormat: "openai",
      models: undefined,
      modelConfigs: undefined,
      contextWindowSize: undefined,
    });
  });
});
