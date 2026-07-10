import {
  DEFAULT_COMPACTION_CONFIG,
  type CompactionConfig,
} from "../../shared/types";
import type { CompactionProvider } from "./compactionProvider";
import {
  DEFAULT_COMPACT_RETRY_CONFIG,
  runAIRequestWithRetry,
  type AIRequestRetryConfig,
} from "./aiRequestRetry";

export function buildCompactRetryConfig(input: {
  compactionConfig?: CompactionConfig;
  retryOverride?: AIRequestRetryConfig;
}): AIRequestRetryConfig {
  const config = input.compactionConfig ?? DEFAULT_COMPACTION_CONFIG;
  const retryCount = Math.max(
    0,
    Math.floor(config.summaryRetryCount ?? DEFAULT_COMPACTION_CONFIG.summaryRetryCount ?? 0)
  );
  return {
    maxRetries: retryCount,
    baseDelayMs: config.summaryRetryBaseDelayMs ?? DEFAULT_COMPACT_RETRY_CONFIG.baseDelayMs,
    maxDelayMs: config.summaryRetryMaxDelayMs ?? DEFAULT_COMPACT_RETRY_CONFIG.maxDelayMs,
    backoffFactor: config.summaryRetryBackoffFactor ?? DEFAULT_COMPACT_RETRY_CONFIG.backoffFactor,
    ...input.retryOverride,
  };
}

export async function generateCompactionSummary(input: {
  provider: CompactionProvider;
  historyPrompt: string;
  compactionConfig?: CompactionConfig;
  retryOverride?: AIRequestRetryConfig;
  signal?: AbortSignal;
}): Promise<string> {
  const config = input.compactionConfig ?? DEFAULT_COMPACTION_CONFIG;
  return runAIRequestWithRetry({
    phase: "compact",
    config: buildCompactRetryConfig({
      compactionConfig: config,
      retryOverride: input.retryOverride,
    }),
    signal: input.signal,
    operation: async () => {
      const summary = await input.provider.generateSummary({
        historyPrompt: input.historyPrompt,
        config,
      });
      const normalized = summary.trim();
      if (!normalized) {
        throw new Error("压缩摘要为空");
      }
      return normalized;
    },
  });
}
