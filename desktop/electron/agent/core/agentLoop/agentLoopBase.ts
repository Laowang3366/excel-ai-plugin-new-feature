import type {
  AgentTurnCallbacks,
  AgentTurnInput,
  CompactionReason,
  Thread,
  ThreadId,
  ThreadRuntimeSnapshot,
} from "../../shared/types";
import type { AIClientConfig } from "../../providers/aiClient";
import { createAIClient } from "../../providers/aiClient";
import { SessionStore } from "../../memory/sessionStore";
import type { StateRuntimeStore } from "../../memory/stateRuntimeStore";
import { enqueueQueuedTurn } from "./queuedTurns";
import {
  createCompactionProvider,
  type CompactionProvider,
} from "./compactionProvider";
import {
  DEFAULT_THREAD_IDLE_UNLOAD_MS,
  ThreadStateManager,
} from "./threadStateManager";
import { PendingInterruptQueue, type ConnectionRequestId } from "./pendingInterruptQueue";
import { ThreadWatchManager } from "./threadWatchManager";
import { InputQueue } from "./inputQueue";
import { TurnState } from "./turnState";
import {
  attachRolloutEventSink as attachRolloutEventSinkHelper,
  bindCallbacksToThread as bindCallbacksToThreadHelper,
} from "./threadRuntime";
import {
  applyAIConfigUpdate,
  mergePendingCompactionReason,
} from "./configUpdates";
import type { AgentLoopConfig } from "./agentLoopConfig";
import type { IdleThreadUnloadTimer } from "./idleThreadUnload";

export abstract class AgentLoopBase {
  protected config: AgentLoopConfig;
  protected aiClient: ReturnType<typeof createAIClient>;
  protected sessionStore: SessionStore;
  protected stateRuntimeStore: StateRuntimeStore | undefined;
  protected readonly turnState = new TurnState();
  protected readonly threadStateManager: ThreadStateManager;
  protected readonly pendingInterruptQueue = new PendingInterruptQueue();
  protected readonly threadWatchManager: ThreadWatchManager;
  protected readonly inputQueue = new InputQueue();
  protected compactionProvider: CompactionProvider;
  protected readonly usesCustomCompactionProvider: boolean;
  protected pendingCompactionReason: CompactionReason | null = null;
  protected idleUnloadTimer: IdleThreadUnloadTimer | null = null;
  protected isDrainingInputQueue = false;
  protected autoDrainInputQueue = true;

  constructor(
    config: AgentLoopConfig,
    sessionStore?: SessionStore,
    stateRuntimeStore?: StateRuntimeStore,
    threadWatchManager = new ThreadWatchManager()
  ) {
    this.config = config;
    this.aiClient = createAIClient(config.aiConfig);
    this.usesCustomCompactionProvider = Boolean(config.compactionProvider);
    this.compactionProvider = config.compactionProvider
      ?? createCompactionProvider(this.aiClient, config.compactionConfig);
    this.sessionStore = sessionStore || new SessionStore();
    this.stateRuntimeStore = stateRuntimeStore;
    this.threadWatchManager = threadWatchManager;
    this.attachRolloutEventSink();
    this.threadStateManager = new ThreadStateManager({
      idleUnloadMs: config.threadIdleUnloadMs ?? DEFAULT_THREAD_IDLE_UNLOAD_MS,
    });
  }

  updateSessionStore(store: SessionStore): void {
    this.sessionStore = store;
    this.attachRolloutEventSink();
  }

  updateStateRuntimeStore(store: StateRuntimeStore): void {
    this.stateRuntimeStore = store;
    this.config.memoryStore?.updateRuntime(store);
    this.attachRolloutEventSink();
  }

  getThread(): Thread | null {
    return this.turnState.activeThread;
  }

  getThreadRuntimeStatus(): ThreadRuntimeSnapshot {
    return this.threadStateManager.getSnapshot();
  }

  getIsRunning(): boolean {
    return this.turnState.isRunning;
  }

  getPendingInterruptRequestIds(): ConnectionRequestId[] {
    return this.pendingInterruptQueue.pendingIds();
  }

  getQueuedInputCount(): number {
    return this.inputQueue.size();
  }

  enqueueTurn(
    input: AgentTurnInput,
    callbacks: AgentTurnCallbacks
  ): { queued: true; queueSize: number } {
    return enqueueQueuedTurn({
      autoDrainInputQueue: this.autoDrainInputQueue,
      inputQueue: this.inputQueue,
      turnInput: input,
      callbacks,
    });
  }

  watchThreadStatus(
    threadId: ThreadId,
    connectionId: string,
    listener: (status: ThreadRuntimeSnapshot) => void
  ) {
    return this.threadWatchManager.watch(threadId, connectionId, listener);
  }

  updateAIConfig(config: AIClientConfig): void {
    const result = applyAIConfigUpdate({
      currentConfig: this.config,
      nextConfig: config,
      activeThread: this.turnState.activeThread,
      usesCustomCompactionProvider: this.usesCustomCompactionProvider,
    });
    this.aiClient = result.aiClient;
    if (result.compactionProvider) this.compactionProvider = result.compactionProvider;
    this.markPendingCompactionReason(result.pendingReason);
  }

  updatePermissionMode(mode: "normal" | "auto_approve_safe" | "confirm_all"): void {
    this.config.permissionMode = mode;
  }

  protected attachRolloutEventSink(): void {
    attachRolloutEventSinkHelper({
      sessionStore: this.sessionStore,
      stateRuntimeStore: this.stateRuntimeStore,
    });
  }

  protected bindCallbacksToThread(
    callbacks: AgentTurnCallbacks,
    threadId: ThreadId,
    clientId?: string
  ): AgentTurnCallbacks {
    return bindCallbacksToThreadHelper({ callbacks, threadId, clientId });
  }

  protected markPendingCompactionReason(reason: CompactionReason | null): void {
    this.pendingCompactionReason = mergePendingCompactionReason(this.pendingCompactionReason, reason);
  }

  protected consumePendingCompactionReason(): CompactionReason | null {
    const reason = this.pendingCompactionReason;
    this.pendingCompactionReason = null;
    return reason;
  }
}
