import type {
  AgentTurnCallbacks,
  CompactProgressItem,
  CompactionConfig,
  CompactionReason,
  Thread,
  ThreadId,
  TurnItem,
} from "../../shared/types";
import type { SessionStore } from "../../memory/sessionStore";
import {
  archiveRolloutIfConfigured,
  completeCompactionProgress,
  failCompactionProgress,
  startCompactionProgress,
} from "./compactionProgress";
import type { CompactionRunnerDependencies } from "./compactionRunner";

export function createCompactionRunnerDeps(input: {
  sessionStore: SessionStore;
  getAllTurnItems: () => TurnItem[];
  generateCompactionSummary: (prompt: string) => Promise<string>;
  getSessionCompactionConfig: () => CompactionConfig;
  archiveRolloutAfterBytes?: number;
  setCompactedHistory: (history: TurnItem[]) => void;
  getActiveThread: () => Thread | null;
  compactionConfig?: CompactionConfig;
}): CompactionRunnerDependencies {
  return {
    sessionStore: input.sessionStore,
    getAllTurnItems: input.getAllTurnItems,
    generateCompactionSummary: input.generateCompactionSummary,
    startCompactionProgress: (
      threadId: ThreadId,
      reason: CompactionReason,
      items: TurnItem[],
      callbacks: AgentTurnCallbacks
    ) => startCompactionProgress({
      sessionStore: input.sessionStore,
      threadId,
      reason,
      items,
      callbacks,
      compactionConfig: input.getSessionCompactionConfig(),
    }),
    completeCompactionProgress: (
      progress: CompactProgressItem,
      tokensBefore: number,
      tokensAfter: number,
      summary: string,
      callbacks: AgentTurnCallbacks
    ) => completeCompactionProgress({
      progress,
      tokensBefore,
      tokensAfter,
      summary,
      callbacks,
    }),
    failCompactionProgress: (
      threadId: ThreadId,
      progress: CompactProgressItem,
      items: TurnItem[],
      error: unknown,
      callbacks: AgentTurnCallbacks
    ) => failCompactionProgress({
      sessionStore: input.sessionStore,
      threadId,
      progress,
      items,
      error,
      callbacks,
    }),
    archiveRolloutIfConfigured: (threadId: ThreadId) => archiveRolloutIfConfigured({
      sessionStore: input.sessionStore,
      threadId,
      threshold: input.archiveRolloutAfterBytes,
    }),
    setCompactedHistory: input.setCompactedHistory,
    getActiveThread: input.getActiveThread,
    compactionConfig: input.compactionConfig,
  };
}
