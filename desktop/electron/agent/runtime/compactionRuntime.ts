import {
  DEFAULT_COMPACTION_CONFIG,
  type CompactionConfig,
} from "../shared/types";

export interface SavedCompactionConfig {
  enabled?: boolean;
  autoCompactThresholdPercent?: number;
  retainedUserMessageMaxTokens?: number;
  retainedRecentItemCount?: number;
  summaryRetryCount?: number;
  summaryRetryBaseDelayMs?: number;
  summaryRetryMaxDelayMs?: number;
  summaryRetryBackoffFactor?: number;
  midTurnThresholdRatio?: number;
  archiveRolloutAfterBytes?: number;
  compactionProvider?: "local" | "remote";
  remoteCompactUrl?: string;
  remoteCompactApiKey?: string;
  remoteCompactModel?: string;
}

/**
 * Agent 压缩配置装配。
 *
 * 关联模块：
 * - core/agentLoop: 消费 CompactionConfig 控制上下文压缩。
 * - main-modules/ipcHandlers: 设置变更时复用同一套阈值计算。
 */
export function buildCompactionConfig(params: {
  contextWindowSize: number;
  savedCompaction?: SavedCompactionConfig | null;
}): CompactionConfig {
  const savedCompaction = params.savedCompaction ?? {};
  const compactionPercent = savedCompaction.autoCompactThresholdPercent ?? 80;

  return {
    enabled: savedCompaction.enabled ?? true,
    autoCompactTokenThreshold: Math.floor(params.contextWindowSize * compactionPercent / 100),
    retainedUserMessageMaxTokens: savedCompaction.retainedUserMessageMaxTokens ?? 20_000,
    retainedRecentItemCount: savedCompaction.retainedRecentItemCount,
    summaryRetryCount: savedCompaction.summaryRetryCount ?? DEFAULT_COMPACTION_CONFIG.summaryRetryCount,
    summaryRetryBaseDelayMs: savedCompaction.summaryRetryBaseDelayMs,
    summaryRetryMaxDelayMs: savedCompaction.summaryRetryMaxDelayMs,
    summaryRetryBackoffFactor: savedCompaction.summaryRetryBackoffFactor,
    midTurnThresholdRatio: savedCompaction.midTurnThresholdRatio ?? 0.9,
    archiveRolloutAfterBytes: savedCompaction.archiveRolloutAfterBytes,
    compactionProvider: savedCompaction.compactionProvider,
    remoteCompactUrl: savedCompaction.remoteCompactUrl,
    remoteCompactApiKey: savedCompaction.remoteCompactApiKey,
    remoteCompactModel: savedCompaction.remoteCompactModel,
    contextWindowSize: params.contextWindowSize,
  };
}
