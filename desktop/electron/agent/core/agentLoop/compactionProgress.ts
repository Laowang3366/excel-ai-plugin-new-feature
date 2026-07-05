import {
  DEFAULT_COMPACTION_CONFIG,
  type AgentTurnCallbacks,
  type CompactProgressItem,
  type CompactionConfig,
  type CompactionReason,
  type ThreadId,
  type TurnItem,
} from "../../shared/types";
import { estimateItemsTokens } from "../../memory/compaction";
import type { SessionStore } from "../../memory/sessionStore";

export async function startCompactionProgress(input: {
  sessionStore: SessionStore;
  threadId: ThreadId;
  reason: CompactionReason;
  items: TurnItem[];
  callbacks: AgentTurnCallbacks;
  compactionConfig: CompactionConfig;
}): Promise<CompactProgressItem> {
  const { sessionStore, threadId, reason, items, callbacks, compactionConfig } = input;
  const tokensBefore = estimateItemsTokens(items);
  const retryCount = Math.max(
    0,
    Math.floor(compactionConfig.summaryRetryCount ?? DEFAULT_COMPACTION_CONFIG.summaryRetryCount ?? 0)
  );
  const timestamp = Date.now();
  const progress: CompactProgressItem = {
    type: "compact_progress",
    id: `compact-progress-${timestamp}-${Math.random().toString(36).slice(2, 8)}`,
    reason,
    status: "running",
    message: "正在压缩上下文...",
    tokensBefore,
    timestamp,
  };

  await sessionStore.appendRolloutItems(threadId, [
    {
      type: "compact_params",
      reason,
      status: "started",
      itemCount: items.length,
      tokensBefore,
    },
  ]);
  callbacks.onEvent({
    type: "thread_compact_started",
    threadId,
    params: {
      reason,
      itemCount: items.length,
      tokensBefore,
      tokenThreshold: compactionConfig.autoCompactTokenThreshold,
      contextWindowSize: compactionConfig.contextWindowSize,
      retryCount,
      timestamp,
    },
  });
  callbacks.onEvent({ type: "item_started", item: progress });
  return progress;
}

export function completeCompactionProgress(input: {
  progress: CompactProgressItem;
  tokensBefore: number;
  tokensAfter: number;
  summary: string;
  callbacks: AgentTurnCallbacks;
}): void {
  const { progress, tokensBefore, tokensAfter, summary, callbacks } = input;
  callbacks.onEvent({
    type: "item_completed",
    item: {
      ...progress,
      status: "completed",
      message: `上下文已压缩：${tokensBefore} → ${tokensAfter} tokens`,
      tokensBefore,
      tokensAfter,
      summary,
      timestamp: Date.now(),
    },
  });
}

export async function failCompactionProgress(input: {
  sessionStore: SessionStore;
  threadId: ThreadId;
  progress: CompactProgressItem;
  items: TurnItem[];
  error: unknown;
  callbacks: AgentTurnCallbacks;
}): Promise<void> {
  const { sessionStore, threadId, progress, items, error, callbacks } = input;
  const message = error instanceof Error ? error.message : String(error);
  await sessionStore.appendRolloutItems(threadId, [
    {
      type: "compact_params",
      reason: progress.reason,
      status: "failed",
      itemCount: items.length,
      tokensBefore: progress.tokensBefore ?? estimateItemsTokens(items),
      error: message,
    },
  ]);
  callbacks.onEvent({
    type: "item_completed",
    item: {
      ...progress,
      status: "failed",
      message: `上下文压缩失败：${message}`,
      timestamp: Date.now(),
    },
  });
}

export async function archiveRolloutIfConfigured(input: {
  sessionStore: SessionStore;
  threadId: ThreadId;
  threshold?: number;
}): Promise<void> {
  if (!input.threshold || input.threshold <= 0) return;

  try {
    await input.sessionStore.spawnRolloutCompressionWorker({
      activeThreadIds: [input.threadId],
      minBytes: input.threshold,
    });
  } catch (error) {
    console.warn("压缩冷 rollout JSONL 失败:", error);
  }
}
