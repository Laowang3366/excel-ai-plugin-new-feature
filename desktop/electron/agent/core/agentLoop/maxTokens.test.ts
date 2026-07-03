import { describe, expect, it } from "vitest";
import { resolveMaxTokens } from "./maxTokens";

describe("resolveMaxTokens", () => {
  it("keeps explicit user maxTokens", () => {
    expect(resolveMaxTokens({
      provider: "openai",
      apiKey: "",
      baseUrl: "",
      model: "gpt-4o",
      maxTokens: 1234,
    })).toBe(1234);
  });

  it("uses context-window based output budget when reasoning is off", () => {
    expect(resolveMaxTokens({
      provider: "openai",
      apiKey: "",
      baseUrl: "",
      model: "gpt-4o",
      contextWindowSize: 128_000,
      reasoningMode: "off",
    })).toBe(7_680);
  });

  it("allocates larger budgets for high and max reasoning", () => {
    expect(resolveMaxTokens({
      provider: "openai",
      apiKey: "",
      baseUrl: "",
      model: "gpt-4o",
      contextWindowSize: 128_000,
      reasoningMode: "high",
    })).toBe(16_384);

    expect(resolveMaxTokens({
      provider: "openai",
      apiKey: "",
      baseUrl: "",
      model: "gpt-4o",
      contextWindowSize: 128_000,
      reasoningMode: "max",
    })).toBe(32_000);
  });

  it("maps enableReasoning without explicit mode to high", () => {
    expect(resolveMaxTokens({
      provider: "openai",
      apiKey: "",
      baseUrl: "",
      model: "gpt-4o",
      contextWindowSize: 64_000,
      enableReasoning: true,
    })).toBe(9_600);
  });
});
