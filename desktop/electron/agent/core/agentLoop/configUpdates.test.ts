import { describe, expect, it, vi } from "vitest";

import type { AIClientConfig } from "../../providers/aiClient";
import type { Thread } from "../../shared/types";
import {
  applyAIConfigUpdate,
  applyCompactionConfigUpdate,
  mergePendingCompactionReason,
} from "./configUpdates";

vi.mock("../../providers/aiClient", async () => {
  const actual = await vi.importActual<typeof import("../../providers/aiClient")>(
    "../../providers/aiClient",
  );
  return {
    ...actual,
    createAIClient: vi.fn(() => ({ streamChat: vi.fn(), chat: vi.fn() })),
  };
});

function createConfig(provider = "openai", model = "gpt-a"): AIClientConfig {
  return {
    provider,
    apiKey: "test",
    baseUrl: "https://example.com/v1",
    model,
    contextWindowSize: 1000,
  };
}

function createThread(): Thread {
  return {
    metadata: {
      threadId: "thread-1",
      preview: "",
      modelProvider: "openai",
      model: "gpt-a",
      contextWindowSize: 1000,
      createdAt: 1,
      updatedAt: 1,
    },
    turns: [],
  };
}

describe("configUpdates", () => {
  it("updates active thread metadata and marks model changes for compaction", () => {
    const currentConfig = {
      aiConfig: createConfig(),
      compactionConfig: {
        enabled: true,
        contextWindowSize: 1000,
        autoCompactTokenThreshold: 800,
        retainedUserMessageMaxTokens: 100,
      },
    };
    const thread = createThread();

    const result = applyAIConfigUpdate({
      currentConfig,
      nextConfig: createConfig("openai", "gpt-b"),
      activeThread: thread,
      usesCustomCompactionProvider: false,
    });

    expect(currentConfig.aiConfig.model).toBe("gpt-b");
    expect(thread.metadata).toMatchObject({
      modelProvider: "openai",
      model: "gpt-b",
      contextWindowSize: 1000,
    });
    expect(thread.metadata.compHash).toBeTruthy();
    expect(result.pendingReason).toBe("model_changed");
    expect(result.compactionProvider).toBeDefined();
  });

  it("updates compaction config and preserves stronger pending reason", () => {
    const currentConfig = {
      compactionConfig: {
        enabled: true,
        contextWindowSize: 1000,
        autoCompactTokenThreshold: 800,
        retainedUserMessageMaxTokens: 100,
      },
    };
    const result = applyCompactionConfigUpdate({
      currentConfig,
      nextConfig: {
        enabled: true,
        contextWindowSize: 500,
        autoCompactTokenThreshold: 400,
        retainedUserMessageMaxTokens: 100,
      },
      aiClient: {} as never,
      activeThread: createThread(),
      usesCustomCompactionProvider: true,
    });

    expect(currentConfig.compactionConfig.contextWindowSize).toBe(500);
    expect(result.compactionProvider).toBeUndefined();
    expect(result.pendingReason).toBe("context_window_changed");
    expect(mergePendingCompactionReason("model_changed", result.pendingReason)).toBe(
      "model_changed",
    );
  });
});
