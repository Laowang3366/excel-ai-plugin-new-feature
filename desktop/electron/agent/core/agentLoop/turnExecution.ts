import type {
  AgentTurnCallbacks,
  AgentTurnInput,
  Thread,
  ThreadId,
  Turn,
} from "../../shared/types";
import type { SessionStore } from "../../memory/sessionStore";
import type { ThreadStateManager } from "./threadStateManager";
import type { TurnState } from "./turnState";
import {
  completeTurn,
  createTurn,
  createUserMessageItem,
} from "./turnRunner";

export function beginTurnRun(turnState: TurnState): void {
  if (turnState.isRunning) {
    throw new Error("Agent 正在运行中，请等待当前 Turn 完成或中断");
  }

  turnState.isRunning = true;
  turnState.abortController = new AbortController();
  turnState.turnCompletionPromise = new Promise((resolve) => {
    turnState.resolveTurnCompletion = resolve;
  });
}

export async function prepareThreadForTurn(input: {
  turnState: TurnState;
  startThread: () => Promise<ThreadId>;
  clearIdleUnloadTimer: () => void;
  threadStateManager: ThreadStateManager;
  publishThreadStatus: () => void;
  persistThreadRuntime: (threadId: ThreadId) => Promise<void>;
  bindCallbacksToThread: (
    callbacks: AgentTurnCallbacks,
    threadId: ThreadId,
    clientId?: string
  ) => AgentTurnCallbacks;
  callbacks: AgentTurnCallbacks;
  clientId?: string;
}): Promise<{ thread: Thread; callbacks: AgentTurnCallbacks }> {
  if (!input.turnState.activeThread) {
    await input.startThread();
  }
  const thread = input.turnState.activeThread!;
  input.clearIdleUnloadTimer();
  input.threadStateManager.markRunning(thread.metadata.threadId);
  input.publishThreadStatus();
  await input.persistThreadRuntime(thread.metadata.threadId);

  return {
    thread,
    callbacks: input.bindCallbacksToThread(input.callbacks, thread.metadata.threadId, input.clientId),
  };
}

export async function createStartedTurn(input: {
  turnInput: AgentTurnInput;
  thread: Thread;
  turnState: TurnState;
  callbacks: AgentTurnCallbacks;
  sessionStore: SessionStore;
  persistThreadSnapshot: (thread: Thread) => Promise<void>;
}): Promise<Turn> {
  const turn = createTurn(input.thread.metadata.threadId);
  input.turnState.activeTurn = turn;
  input.thread.metadata.activeTurnId = turn.turnId;
  input.thread.metadata.lastTurnStatus = "in_progress";
  await input.persistThreadSnapshot(input.thread);

  input.callbacks.onEvent({ type: "turn_started", turnId: turn.turnId });

  const userItem = createUserMessageItem(input.turnInput);
  turn.items.push(userItem);
  await input.sessionStore.appendTurnItem(input.thread.metadata.threadId, turn.turnId, userItem);
  input.callbacks.onEvent({ type: "item_started", item: userItem });
  input.callbacks.onEvent({ type: "item_completed", item: userItem });

  if (!input.thread.metadata.preview) {
    input.thread.metadata.preview = input.turnInput.content.slice(0, 100);
  }

  return turn;
}

export async function completeSuccessfulTurn(input: {
  thread: Thread;
  turn: Turn;
  callbacks: AgentTurnCallbacks;
  sessionStore: SessionStore;
  persistThreadSnapshot: (thread: Thread) => Promise<void>;
  scheduleTurnMemoryExtraction: (thread: Thread, turn: Turn) => void;
}): Promise<Turn> {
  completeTurn(input.turn);
  if (input.turn.tokenUsage) {
    await input.sessionStore.appendTurnUsage(
      input.thread.metadata.threadId,
      input.turn.turnId,
      input.turn.tokenUsage
    );
  }
  input.thread.turns.push(input.turn);
  input.thread.metadata.updatedAt = Date.now();
  input.thread.metadata.lastTurnStatus = input.turn.status;
  input.thread.metadata.activeTurnId = undefined;
  await input.persistThreadSnapshot(input.thread);

  input.callbacks.onEvent({
    type: "turn_completed",
    turnId: input.turn.turnId,
    usage: input.turn.tokenUsage,
  });
  input.scheduleTurnMemoryExtraction(input.thread, input.turn);
  return input.turn;
}

export async function handleTurnFailure(input: {
  error: unknown;
  turnState: TurnState;
  callbacks: AgentTurnCallbacks;
  persistThreadSnapshot: (thread: Thread) => Promise<void>;
}): Promise<void> {
  const err = input.error as { name?: string; message?: string };
  const message = err.message ?? String(input.error);
  const activeTurn = input.turnState.activeTurn;
  if (!activeTurn) return;

  activeTurn.status = err.name === "AbortError" ? "interrupted" : "failed";
  activeTurn.error = message;
  activeTurn.completedAt = Date.now();

  if (activeTurn.status === "interrupted") {
    input.callbacks.onEvent({ type: "turn_interrupted", turnId: activeTurn.turnId });
  } else {
    input.callbacks.onEvent({
      type: "turn_failed",
      turnId: activeTurn.turnId,
      error: message,
    });
  }

  input.turnState.activeThread?.turns.push(activeTurn);
  if (!input.turnState.activeThread) return;

  input.turnState.activeThread.metadata.updatedAt = activeTurn.completedAt ?? Date.now();
  input.turnState.activeThread.metadata.lastTurnStatus = activeTurn.status;
  input.turnState.activeThread.metadata.activeTurnId = undefined;
  await input.persistThreadSnapshot(input.turnState.activeThread);
}

export async function finishTurnRun(input: {
  turnState: TurnState;
  threadStateManager: ThreadStateManager;
  publishThreadStatus: () => void;
  scheduleIdleThreadUnload: () => void;
  persistThreadRuntime: (threadId: ThreadId) => Promise<void>;
}): Promise<void> {
  input.turnState.isRunning = false;
  input.turnState.abortController = null;

  if (input.turnState.activeThread) {
    input.threadStateManager.markIdle(input.turnState.activeThread.metadata.threadId);
    input.publishThreadStatus();
    input.scheduleIdleThreadUnload();
    await input.persistThreadRuntime(input.turnState.activeThread.metadata.threadId);
  }

  input.turnState.resolveTurnCompletion?.();
  input.turnState.turnCompletionPromise = null;
  input.turnState.resolveTurnCompletion = null;
}
