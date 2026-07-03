import { describe, expect, it } from "vitest";

import { buildCompactionConfig } from "./compactionRuntime";

describe("buildCompactionConfig", () => {
  it("maps saved health and archive settings into runtime config", () => {
    const config = buildCompactionConfig({
      contextWindowSize: 1000,
      savedCompaction: {
        enabled: true,
        autoCompactThresholdPercent: 75,
        retainedUserMessageMaxTokens: 120,
        retainedRecentItemCount: 3,
        summaryRetryCount: 2,
        summaryRetryBaseDelayMs: 100,
        summaryRetryMaxDelayMs: 500,
        summaryRetryBackoffFactor: 3,
        midTurnThresholdRatio: 0.8,
        archiveRolloutAfterBytes: 4096,
        compactionProvider: "remote",
        remoteCompactUrl: "https://compact.example.test/v2",
        remoteCompactApiKey: "remote-key",
        remoteCompactModel: "compact-model",
      },
    });

    expect(config).toMatchObject({
      enabled: true,
      autoCompactTokenThreshold: 750,
      retainedUserMessageMaxTokens: 120,
      retainedRecentItemCount: 3,
      summaryRetryCount: 2,
      summaryRetryBaseDelayMs: 100,
      summaryRetryMaxDelayMs: 500,
      summaryRetryBackoffFactor: 3,
      midTurnThresholdRatio: 0.8,
      archiveRolloutAfterBytes: 4096,
      contextWindowSize: 1000,
      compactionProvider: "remote",
      remoteCompactUrl: "https://compact.example.test/v2",
      remoteCompactApiKey: "remote-key",
      remoteCompactModel: "compact-model",
    });
  });
});
