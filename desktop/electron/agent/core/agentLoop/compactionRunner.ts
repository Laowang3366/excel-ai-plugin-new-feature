import type {
  AgentTurnCallbacks,
  CompactProgressItem,
  CompactionConfig,
  CompactionReason,
  Thread,
  ThreadId,
  Turn,
  TurnItem,
} from "../../shared/types";
import {
  historyToCompactPrompt,
  performCompaction,
} from "../../memory/compaction";
import type { SessionStore } from "../../memory/sessionStore";

export interface CompactionRunnerDependencies {
  sessionStore: SessionStore;
  getAllTurnItems: () => TurnItem[];
  generateCompactionSummary: (prompt: string) => Promise<string>;
  startCompactionProgress: (
    threadId: ThreadId,
    reason: CompactionReason,
    items: TurnItem[],
    callbacks: AgentTurnCallbacks
  ) => Promise<CompactProgressItem>;
  completeCompactionProgress: (
    progress: CompactProgressItem,
    tokensBefore: number,
    tokensAfter: number,
    summary: string,
    callbacks: AgentTurnCallbacks
  ) => void;
  failCompactionProgress: (
    threadId: ThreadId,
    progress: CompactProgressItem,
    items: TurnItem[],
    error: unknown,
    callbacks: AgentTurnCallbacks
  ) => Promise<void>;
  archiveRolloutIfConfigured: (threadId: ThreadId) => Promise<void>;
  setCompactedHistory: (history: TurnItem[]) => void;
  getActiveThread: () => Thread | null;
  compactionConfig?: CompactionConfig;
}

export async function runAutoCompaction(input: {
  thread: Thread;
  reason: CompactionReason;
  callbacks: AgentTurnCallbacks;
  deps: CompactionRunnerDependencies;
}): Promise<void> {
  const { thread, reason, callbacks, deps } = input;
  const allItems = deps.getAllTurnItems();
  if (allItems.length === 0) return;

  const prompt = historyToCompactPrompt(allItems);
  const progress = await deps.startCompactionProgress(
    thread.metadata.threadId,
    reason,
    allItems,
    callbacks
  );
  let summary: string;
  try {
    summary = await deps.generateCompactionSummary(prompt);
  } catch (error) {
    await deps.failCompactionProgress(thread.metadata.threadId, progress, allItems, error, callbacks);
    throw error;
  }

  const { compactedItem, newHistory } = performCompaction(
    allItems,
    summary,
    reason,
    deps.compactionConfig
  );
  deps.setCompactedHistory(newHistory);
  thread.turns = [];

  await writeCompletedCompaction({
    sessionStore: deps.sessionStore,
    threadId: thread.metadata.threadId,
    summary,
    replacementHistory: newHistory,
    reason,
    itemCount: allItems.length,
    tokensBefore: compactedItem.tokensBefore,
    tokensAfter: compactedItem.tokensAfter,
  });
  await deps.archiveRolloutIfConfigured(thread.metadata.threadId);
  emitCompletedCompaction({
    callbacks,
    completeCompactionProgress: deps.completeCompactionProgress,
    progress,
    summary,
    tokensBefore: compactedItem.tokensBefore,
    tokensAfter: compactedItem.tokensAfter,
  });
}

export async function runMidTurnCompaction(input: {
  turn: Turn;
  callbacks: AgentTurnCallbacks;
  deps: CompactionRunnerDependencies;
}): Promise<void> {
  const { turn, callbacks, deps } = input;
  const currentUserItems = turn.items.filter((item) => item.type === "user_message");
  const currentUserItemIds = new Set(currentUserItems.map((item) => item.id));
  const allItems = deps.getAllTurnItems();
  const prompt = historyToCompactPrompt(allItems);
  const activeThread = deps.getActiveThread();
  const threadId = activeThread?.metadata.threadId;
  const progress = threadId
    ? await deps.startCompactionProgress(threadId, "auto_token_limit", allItems, callbacks)
    : null;
  let summary: string;
  try {
    summary = await deps.generateCompactionSummary(prompt);
  } catch (error) {
    if (threadId && progress) {
      await deps.failCompactionProgress(threadId, progress, allItems, error, callbacks);
    }
    throw error;
  }

  const { compactedItem, newHistory } = performCompaction(
    allItems,
    summary,
    "auto_token_limit",
    deps.compactionConfig
  );
  deps.setCompactedHistory(
    newHistory.filter((item) => item.type !== "user_message" || !currentUserItemIds.has(item.id))
  );
  turn.items = currentUserItems;

  const latestActiveThread = deps.getActiveThread();
  if (latestActiveThread) {
    latestActiveThread.turns = [];
    await writeCompletedCompaction({
      sessionStore: deps.sessionStore,
      threadId: latestActiveThread.metadata.threadId,
      summary,
      replacementHistory: newHistory,
      reason: "auto_token_limit",
      itemCount: allItems.length,
      tokensBefore: compactedItem.tokensBefore,
      tokensAfter: compactedItem.tokensAfter,
    });
    await deps.archiveRolloutIfConfigured(latestActiveThread.metadata.threadId);
  }

  if (progress) {
    emitCompletedCompaction({
      callbacks,
      completeCompactionProgress: deps.completeCompactionProgress,
      progress,
      summary,
      tokensBefore: compactedItem.tokensBefore,
      tokensAfter: compactedItem.tokensAfter,
    });
  } else {
    callbacks.onEvent({
      type: "context_compacted",
      summary,
      tokensBefore: compactedItem.tokensBefore,
      tokensAfter: compactedItem.tokensAfter,
    });
  }
}

async function writeCompletedCompaction(input: {
  sessionStore: SessionStore;
  threadId: ThreadId;
  summary: string;
  replacementHistory: TurnItem[];
  reason: CompactionReason;
  itemCount: number;
  tokensBefore: number;
  tokensAfter: number;
}): Promise<void> {
  await input.sessionStore.appendRolloutItems(input.threadId, [
    {
      type: "compacted",
      summary: input.summary,
      replacementHistory: input.replacementHistory,
    },
    {
      type: "compact_params",
      reason: input.reason,
      status: "completed",
      itemCount: input.itemCount,
      tokensBefore: input.tokensBefore,
      tokensAfter: input.tokensAfter,
    },
  ]);
}

function emitCompletedCompaction(input: {
  callbacks: AgentTurnCallbacks;
  completeCompactionProgress: CompactionRunnerDependencies["completeCompactionProgress"];
  progress: CompactProgressItem;
  summary: string;
  tokensBefore: number;
  tokensAfter: number;
}): void {
  input.completeCompactionProgress(
    input.progress,
    input.tokensBefore,
    input.tokensAfter,
    input.summary,
    input.callbacks
  );
  input.callbacks.onEvent({
    type: "context_compacted",
    summary: input.summary,
    tokensBefore: input.tokensBefore,
    tokensAfter: input.tokensAfter,
  });
}
