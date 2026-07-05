import type { AIClientConfig } from "../../providers/aiClient";
import type {
  CompactionConfig,
  Thread,
  ThreadId,
  Turn,
  TurnItem,
} from "../../shared/types";
import type { SessionStore } from "../../memory/sessionStore";
import type { ThreadStateManager } from "./threadStateManager";
import type { TurnState } from "./turnState";
import { createAgentThread, loadAgentThread } from "./threadLifecycle";

type PersistThreadSnapshot = (thread: Thread) => Promise<void>;
type PersistThreadRuntime = (threadId: ThreadId) => Promise<void>;

export async function resetThreadSession(input: {
  isRunning: boolean;
  interrupt: () => Promise<void>;
  clearIdleUnloadTimer: () => void;
  turnState: TurnState;
  threadStateManager: ThreadStateManager;
  folderId?: string;
}): Promise<void> {
  if (input.isRunning) {
    await input.interrupt();
  }
  input.clearIdleUnloadTimer();
  input.turnState.resetForNextThread(input.folderId);
  input.threadStateManager.clear();
}

export async function startThreadSession(input: {
  turnState: TurnState;
  sessionStore: SessionStore;
  aiConfig: AIClientConfig;
  compactionConfig?: CompactionConfig;
  setActiveThread: (thread: Thread | null) => void;
  setCompactedHistory: (history: TurnItem[] | null) => void;
  threadStateManager: ThreadStateManager;
  publishThreadStatus: () => void;
  scheduleIdleThreadUnload: () => void;
  persistThreadSnapshot: PersistThreadSnapshot;
  persistThreadRuntime: PersistThreadRuntime;
}): Promise<ThreadId> {
  const folderId = input.turnState.consumePendingFolderId();
  const thread = await createAgentThread({
    sessionStore: input.sessionStore,
    aiConfig: input.aiConfig,
    compactionConfig: input.compactionConfig,
    folderId,
  });

  input.setActiveThread(thread);
  input.setCompactedHistory(null);
  input.threadStateManager.markLoaded(thread.metadata.threadId);
  input.publishThreadStatus();
  input.scheduleIdleThreadUnload();
  await input.persistThreadSnapshot(thread);
  await input.persistThreadRuntime(thread.metadata.threadId);
  return thread.metadata.threadId;
}

export async function resumeThreadSession(input: {
  isRunning: boolean;
  activeThread: Thread | null;
  sessionStore: SessionStore;
  threadId: ThreadId;
  setActiveThread: (thread: Thread | null) => void;
  setCompactedHistory: (history: TurnItem[] | null) => void;
  threadStateManager: ThreadStateManager;
  publishThreadStatus: () => void;
  scheduleIdleThreadUnload: () => void;
  persistThreadSnapshot: PersistThreadSnapshot;
  persistThreadRuntime: PersistThreadRuntime;
}): Promise<boolean> {
  if (input.isRunning) {
    return input.activeThread?.metadata.threadId === input.threadId;
  }

  const result = await loadAgentThread(input.sessionStore, input.threadId);
  if (!result) return false;

  input.setActiveThread(result.thread);
  input.setCompactedHistory(result.compactedHistory);
  input.threadStateManager.markLoaded(result.thread.metadata.threadId);
  input.publishThreadStatus();
  input.scheduleIdleThreadUnload();
  await input.persistThreadSnapshot(result.thread);
  await input.persistThreadRuntime(result.thread.metadata.threadId);
  return true;
}

export async function sweepIdleThreadSession(input: {
  now: number;
  isRunning: boolean;
  activeThread: Thread | null;
  sessionStore: SessionStore;
  setActiveThread: (thread: Thread | null) => void;
  setActiveTurn: (turn: Turn | null) => void;
  setCompactedHistory: (history: TurnItem[] | null) => void;
  threadStateManager: ThreadStateManager;
  publishThreadStatus: () => void;
  clearIdleUnloadTimer: () => void;
  persistThreadRuntime: PersistThreadRuntime;
}): Promise<boolean> {
  if (input.isRunning || !input.activeThread || !input.threadStateManager.shouldUnload(input.now)) {
    return false;
  }

  const threadId = input.activeThread.metadata.threadId;
  await input.sessionStore.flushRolloutWrites();
  input.setActiveThread(null);
  input.setActiveTurn(null);
  input.setCompactedHistory(null);
  input.threadStateManager.markUnloaded(input.now);
  input.publishThreadStatus();
  input.clearIdleUnloadTimer();
  await input.persistThreadRuntime(threadId);
  return true;
}
