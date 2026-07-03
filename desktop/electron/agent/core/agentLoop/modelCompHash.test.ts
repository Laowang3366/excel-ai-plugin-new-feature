import { describe, expect, it } from "vitest";

import { isModelCompHashCompatible, resolveModelCompHash } from "./modelCompHash";

describe("modelCompHash", () => {
  it("uses explicit compHash when configured", () => {
    expect(resolveModelCompHash({
      provider: "openai",
      apiKey: "",
      baseUrl: "",
      model: "model-a",
      compHash: "responses-family",
    })).toBe("responses-family");
  });

  it("falls back to api format, provider, and model for safe defaults", () => {
    expect(resolveModelCompHash({
      provider: "openai",
      apiKey: "",
      baseUrl: "",
      model: "model-a",
      apiFormat: "responses",
    })).toBe("responses:openai:model-a");
  });

  it("treats models with the same explicit compHash as compatible", () => {
    expect(isModelCompHashCompatible(
      { provider: "openai", apiKey: "", baseUrl: "", model: "model-a", compHash: "office-chat-v1" },
      { provider: "custom", apiKey: "", baseUrl: "", model: "model-b", compHash: "office-chat-v1" }
    )).toBe(true);
  });

  it("treats models with different fallback hashes as incompatible", () => {
    expect(isModelCompHashCompatible(
      { provider: "openai", apiKey: "", baseUrl: "", model: "model-a", apiFormat: "openai" },
      { provider: "openai", apiKey: "", baseUrl: "", model: "model-b", apiFormat: "openai" }
    )).toBe(false);
  });
});
