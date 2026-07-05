import { describe, expect, it, vi } from "vitest";

import {
  buildCompactRetryConfig,
  generateCompactionSummary,
} from "./compactionSummary";

describe("compactionSummary", () => {
  it("builds retry config from compaction settings and explicit overrides", () => {
    expect(buildCompactRetryConfig({
      compactionConfig: {
        enabled: true,
        contextWindowSize: 1000,
        autoCompactTokenThreshold: 800,
        retainedUserMessageMaxTokens: 100,
        summaryRetryCount: 2,
        summaryRetryBaseDelayMs: 10,
        summaryRetryMaxDelayMs: 20,
        summaryRetryBackoffFactor: 3,
      },
      retryOverride: {
        maxRetries: 5,
        baseDelayMs: 1,
      },
    })).toMatchObject({
      maxRetries: 5,
      baseDelayMs: 1,
      maxDelayMs: 20,
      backoffFactor: 3,
    });
  });

  it("generates summary through provider with normalized config", async () => {
    const provider = {
      generateSummary: vi.fn().mockResolvedValue("摘要"),
    };

    await expect(generateCompactionSummary({
      provider,
      historyPrompt: "历史",
      compactionConfig: {
        enabled: true,
        contextWindowSize: 1000,
        autoCompactTokenThreshold: 800,
        retainedUserMessageMaxTokens: 100,
        summaryRetryCount: 0,
      },
    })).resolves.toBe("摘要");

    expect(provider.generateSummary).toHaveBeenCalledWith({
      historyPrompt: "历史",
      config: expect.objectContaining({
        enabled: true,
        contextWindowSize: 1000,
      }),
    });
  });
});
