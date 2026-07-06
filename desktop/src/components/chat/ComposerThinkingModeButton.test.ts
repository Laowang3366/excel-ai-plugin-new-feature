import { describe, expect, it } from "vitest";
import type { AiProviderConfig } from "../../electronApi";
import type { ProviderTemplate } from "../../store/settingsProviderTemplates";
import { resolveComposerThinkingModeState } from "./ComposerThinkingModeButton";

const templates: ProviderTemplate[] = [
  {
    id: "openai",
    name: "OpenAI",
    provider: "openai",
    baseUrl: "https://api.openai.com/v1",
    defaultModel: "gpt-5",
    defaultContextWindowSize: 256000,
    apiFormat: "responses",
    category: "direct",
    reasoningOptions: [{ value: "off", label: "Off" }, { value: "medium", label: "Medium" }],
    defaultReasoningMode: "medium",
  },
];

function provider(patch: Partial<AiProviderConfig> = {}): AiProviderConfig {
  return {
    id: "p1",
    name: "OpenAI",
    provider: "openai",
    apiKey: "sk",
    baseUrl: "https://api.openai.com/v1",
    model: "gpt-5",
    apiFormat: "responses",
    reasoningMode: "high",
    ...patch,
  };
}

describe("resolveComposerThinkingModeState", () => {
  it("returns null when the active provider is missing", () => {
    expect(resolveComposerThinkingModeState({
      providers: {},
      activeProviderId: "missing",
      templates,
      language: "en-US",
    })).toBeNull();
  });

  it("uses per-model reasoning mode before provider-level mode", () => {
    const state = resolveComposerThinkingModeState({
      providers: {
        p1: provider({
          modelConfigs: [{ name: "gpt-5", reasoningMode: "low" }],
        }),
      },
      activeProviderId: "p1",
      templates,
      language: "en-US",
    });

    expect(state?.currentMode).toBe("low");
    expect(state?.isReasoningActive).toBe(true);
    expect(state?.options.map((option) => option.value)).toContain("max");
  });

  it("falls back to the adapted default when configured mode is invalid", () => {
    const state = resolveComposerThinkingModeState({
      providers: {
        p1: provider({ reasoningMode: undefined }),
      },
      activeProviderId: "p1",
      templates,
      language: "en-US",
    });

    expect(state?.currentMode).toBe("medium");
  });
});
