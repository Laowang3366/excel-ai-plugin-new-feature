import { describe, expect, it } from "vitest";
import type { AiProviderConfig } from "../electronApi";
import {
  buildReasoningOptions,
  coerceReasoningMode,
  defaultReasoningModeForOptions,
  formatReasoningOptionLabels,
  normalizeProviderReasoningConfig,
  REASONING_FULL,
  REASONING_HIGH_MAX,
  REASONING_TOGGLE,
  resolveReasoningOptionValues,
} from "./reasoningSupport";

describe("reasoningSupport", () => {
  it("enables the full reasoning scale for OpenAI responses-compatible models", () => {
    expect(resolveReasoningOptionValues({
      provider: "custom",
      apiFormat: "responses",
      model: "gpt-5",
    })).toEqual(REASONING_FULL);

    expect(resolveReasoningOptionValues({
      provider: "openai",
      apiFormat: "chat",
      model: "codex-mini",
    })).toEqual(REASONING_FULL);
  });

  it("maps known reasoning providers to their supported option sets", () => {
    expect(resolveReasoningOptionValues({
      provider: "aggregator",
      apiFormat: "chat",
      model: "qwen3-max",
    })).toEqual(REASONING_HIGH_MAX);

    expect(resolveReasoningOptionValues({
      provider: "moonshot",
      apiFormat: "chat",
      model: "kimi-k2",
    })).toEqual(REASONING_TOGGLE);
  });

  it("deduplicates template modes and falls back to toggle options when template data is too weak", () => {
    expect(resolveReasoningOptionValues(
      { provider: "custom", apiFormat: "chat", model: "plain" },
      {
        reasoningOptions: [
          { value: "high", label: "High" },
          { value: "high", label: "Duplicate" },
          { value: "max", label: "Max" },
          { value: "invalid", label: "Invalid" },
        ],
      },
    )).toEqual(["high", "max"]);

    expect(resolveReasoningOptionValues(
      { provider: "custom", apiFormat: "chat", model: "plain" },
      { reasoningOptions: [{ value: "invalid", label: "Invalid" }] },
    )).toEqual(REASONING_TOGGLE);
  });

  it("coerces invalid selected modes to the best supported default", () => {
    expect(defaultReasoningModeForOptions(["off", "low", "medium"])).toBe("medium");
    expect(defaultReasoningModeForOptions(["off", "high"], "high")).toBe("high");
    expect(coerceReasoningMode("low", ["off", "high"], "high")).toBe("high");
  });

  it("formats localized option labels", () => {
    expect(buildReasoningOptions(["off", "high"], "en-US")).toEqual([
      { value: "off", label: "Off" },
      { value: "high", label: "High" },
    ]);
    expect(formatReasoningOptionLabels(["off", "high"], "zh-CN")).toBe("关闭 / 高");
  });

  it("normalizes provider and per-model reasoning modes against supported options", () => {
    const provider = {
      provider: "custom",
      apiFormat: "chat",
      model: "qwen3-max",
      reasoningMode: "low",
      modelConfigs: [
        { name: "qwen3-max", reasoningMode: "max", reasoningOptions: ["legacy"] },
        { name: "kimi-k2", reasoningMode: "medium", reasoningOptions: ["legacy"] },
      ],
    } as AiProviderConfig;

    const normalized = normalizeProviderReasoningConfig(provider);

    expect(normalized.reasoningMode).toBe("max");
    expect(normalized.enableReasoning).toBe(true);
    expect(normalized.modelConfigs).toEqual([
      { name: "qwen3-max", reasoningMode: "max" },
      { name: "kimi-k2" },
    ]);
  });
});
