import { describe, expect, it } from "vitest";
import type { AiProviderConfig } from "../../electronApi";
import { buildEditProviderPatch } from "./editProviderPatch";

const baseProvider: AiProviderConfig = {
  id: "provider-1",
  name: "Test Provider",
  provider: "custom",
  apiKey: "key-1",
  baseUrl: "https://api.example.com/v1",
  model: "model-a",
  apiFormat: "openai",
  contextWindowSize: 128000,
  reasoningMode: "off",
  modelConfigs: [
    {
      name: "model-a",
      contextWindowSize: 128000,
      reasoningMode: "off",
    },
  ],
};

describe("buildEditProviderPatch", () => {
  it("returns an empty patch when the editable draft matches the provider", () => {
    expect(buildEditProviderPatch(baseProvider, {
      name: "Test Provider",
      apiFormat: "openai",
      baseUrl: "https://api.example.com/v1",
      apiKey: "key-1",
      model: "model-a",
      contextWindowSize: 128000,
      effectiveReasoningMode: "off",
      modelConfigs: [
        {
          name: "model-a",
          contextWindowSize: 128000,
          reasoningMode: "off",
          reasoningOptions: ["off", "low"],
        },
      ],
    })).toEqual({});
  });

  it("includes changed fields and strips legacy model reasoning options", () => {
    expect(buildEditProviderPatch(baseProvider, {
      name: "Renamed Provider",
      apiFormat: "responses",
      baseUrl: "https://api.new.example/v1",
      apiKey: "key-2",
      model: "model-b",
      contextWindowSize: 256000,
      effectiveReasoningMode: "medium",
      modelConfigs: [
        {
          name: "model-b",
          contextWindowSize: 256000,
          reasoningMode: "medium",
          reasoningOptions: ["medium", "high"],
        },
      ],
    })).toEqual({
      name: "Renamed Provider",
      apiFormat: "responses",
      baseUrl: "https://api.new.example/v1",
      apiKey: "key-2",
      model: "model-b",
      contextWindowSize: 256000,
      reasoningMode: "medium",
      modelConfigs: [
        {
          name: "model-b",
          contextWindowSize: 256000,
          reasoningMode: "medium",
        },
      ],
    });
  });
});
