import type { SessionStore } from "../../memory/sessionStore";
import type {
  AgentTurnCallbacks,
  AgentTurnInput,
  CompactionConfig,
  CompactionReason,
  Thread,
  ThreadId,
  Turn,
  TurnItem,
} from "../../shared/types";
import type { ThreadStateManager } from "./threadStateManager";
import type { TurnState } from "./turnState";
import { buildPreTurnCompactionPlan } from "./preTurnCompaction";
import {
  beginTurnRun,
  completeSuccessfulTurn,
  createStartedTurn,
  finishTurnRun,
  handleTurnFailure,
  prepareThreadForTurn,
} from "./turnExecution";

export async function runTurnFlow(input: {
  turnInput: AgentTurnInput;
  callbacks: AgentTurnCallbacks;
  turnState: TurnState;
  sessionStore: SessionStore;
  threadStateManager: ThreadStateManager;
  setAutoDrainInputQueue: (enabled: boolean) => void;
  shouldDrainInputQueue: () => boolean;
  scheduleInputQueueDrain: () => void;
  startThread: () => Promise<ThreadId>;
  clearIdleUnloadTimer: () => void;
  publishThreadStatus: () => void;
  persistThreadRuntime: (threadId: ThreadId) => Promise<void>;
  bindCallbacksToThread: (
    callbacks: AgentTurnCallbacks,
    threadId: ThreadId,
    clientId?: string
  ) => AgentTurnCallbacks;
  getAllTurnItems: () => TurnItem[];
  compactionConfig?: CompactionConfig;
  consumePendingCompactionReason: () => CompactionReason | null;
  performAutoCompaction: (
    thread: Thread,
    reason: CompactionReason,
    callbacks: AgentTurnCallbacks
  ) => Promise<void>;
  persistThreadSnapshot: (thread: Thread) => Promise<void>;
  runAgentLoop: (
    turn: Turn,
    callbacks: AgentTurnCallbacks,
    turnInput: AgentTurnInput,
    resumeContext?: string
  ) => Promise<void>;
  scheduleTurnMemoryExtraction: (thread: Thread, turn: Turn) => void;
  scheduleIdleThreadUnload: () => void;
}): Promise<Turn> {
  beginTurnRun(input.turnState);
  input.setAutoDrainInputQueue(true);

  let turnCallbacks = input.callbacks;

  try {
    const prepared = await prepareThreadForTurn({
      turnState: input.turnState,
      startThread: input.startThread,
      clearIdleUnloadTimer: input.clearIdleUnloadTimer,
      threadStateManager: input.threadStateManager,
      publishThreadStatus: input.publishThreadStatus,
      persistThreadRuntime: input.persistThreadRuntime,
      bindCallbacksToThread: input.bindCallbacksToThread,
      callbacks: input.callbacks,
      clientId: input.turnInput.clientId,
    });
    const { thread } = prepared;
    turnCallbacks = prepared.callbacks;

    const allItems = input.getAllTurnItems();
    const compactionPlan = buildPreTurnCompactionPlan({
      items: allItems,
      thread,
      globalConfig: input.compactionConfig,
      pendingReason: input.consumePendingCompactionReason(),
    });
    if (compactionPlan.reason) {
      await input.performAutoCompaction(thread, compactionPlan.reason, turnCallbacks);
    }

    const turn = await createStartedTurn({
      turnInput: input.turnInput,
      thread,
      turnState: input.turnState,
      callbacks: turnCallbacks,
      sessionStore: input.sessionStore,
      persistThreadSnapshot: input.persistThreadSnapshot,
    });

    const resumeContext = input.turnInput.isResume ? input.turnInput.resumeContext : undefined;
    await input.runAgentLoop(turn, turnCallbacks, input.turnInput, resumeContext);

    return completeSuccessfulTurn({
      thread,
      turn,
      callbacks: turnCallbacks,
      sessionStore: input.sessionStore,
      persistThreadSnapshot: input.persistThreadSnapshot,
      scheduleTurnMemoryExtraction: input.scheduleTurnMemoryExtraction,
    });
  } catch (err: any) {
    await handleTurnFailure({
      error: err,
      turnState: input.turnState,
      callbacks: turnCallbacks,
      persistThreadSnapshot: input.persistThreadSnapshot,
    });
    throw err;
  } finally {
    await finishTurnRun({
      turnState: input.turnState,
      threadStateManager: input.threadStateManager,
      publishThreadStatus: input.publishThreadStatus,
      scheduleIdleThreadUnload: input.scheduleIdleThreadUnload,
      persistThreadRuntime: input.persistThreadRuntime,
    });
    if (input.shouldDrainInputQueue()) {
      input.scheduleInputQueueDrain();
    }
  }
}
