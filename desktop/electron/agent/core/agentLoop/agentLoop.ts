import {
  type TurnItem,
  type Turn,
  type Thread,
  type ThreadId,
  type TokenUsage,
  type AgentTurnInput,
  type AgentTurnCallbacks,
  type ToolDefinition,
  type CompactionConfig,
  type CompactionReason,
  DEFAULT_COMPACTION_CONFIG,
} from "../../shared/types";

import {
  buildResumeContext,
} from "../../memory/compaction";

import { getToolDefinitions as getToolDefs } from "./toolExecutor";
import { buildSessionCompactionConfig } from "./sessionCompactionConfig";
import type { ConnectionRequestId } from "./pendingInterruptQueue";
import {
  buildContextUsageEvent,
  collectPromptTurnItemGroups,
  collectPromptTurnItems,
} from "./contextUsage";
import {
  persistThreadRuntime as persistThreadRuntimeHelper,
  persistThreadSnapshot as persistThreadSnapshotHelper,
  scheduleTurnMemoryExtraction as scheduleTurnMemoryExtractionHelper,
} from "./threadRuntime";
import {
  runAutoCompaction,
  runMidTurnCompaction,
} from "./compactionRunner";
import {
  resetThreadSession,
  resumeThreadSession,
  startThreadSession,
  sweepIdleThreadSession,
} from "./threadSession";
import {
  drainQueuedTurnsAndReschedule,
  interruptCurrentTurn,
  scheduleQueuedTurnsDrain,
} from "./queuedTurns";
import {
  applyCompactionConfigUpdate,
} from "./configUpdates";
import { generateCompactionSummary as generateCompactionSummaryWithRetry } from "./compactionSummary";
import { createCompactionRunnerDeps } from "./compactionRunnerDeps";
import {
  clearIdleThreadUnloadTimer,
  scheduleIdleThreadUnload as scheduleIdleThreadUnloadTimer,
} from "./idleThreadUnload";
import { runTurnFlow } from "./turnFlow";
import { runAgentLoopWithDeps } from "./agentLoopRoundDeps";
import { AgentLoopBase } from "./agentLoopBase";
import { createLogger } from "../../../shared/logger";

export type { AgentLoopConfig } from "./agentLoopConfig";

const agentLoopLogger = createLogger("AgentLoop");

export class AgentLoop extends AgentLoopBase {

  private get threadSessionDeps() {
    return {
      turnState: this.turnState,
      sessionStore: this.sessionStore,
      threadStateManager: this.threadStateManager,
      setActiveThread: (thread: Thread | null) => { this.turnState.activeThread = thread; },
      setActiveTurn: (turn: Turn | null) => { this.turnState.activeTurn = turn; },
      setCompactedHistory: (history: TurnItem[] | null) => { this.turnState.compactedHistory = history; },
      publishThreadStatus: () => this.publishThreadStatus(),
      scheduleIdleThreadUnload: () => this.scheduleIdleThreadUnload(),
      clearIdleUnloadTimer: () => this.clearIdleUnloadTimer(),
      persistThreadSnapshot: (thread: Thread) => this.persistThreadSnapshot(thread),
      persistThreadRuntime: (threadId: ThreadId) => this.persistThreadRuntime(threadId),
    };
  }

  async resetThread(folderId?: string): Promise<void> {
    await resetThreadSession({
      ...this.threadSessionDeps,
      isRunning: this.turnState.isRunning,
      interrupt: () => this.interrupt(),
      folderId,
    });
  }

  async startThread(): Promise<ThreadId> {
    return startThreadSession({
      ...this.threadSessionDeps,
      aiConfig: this.config.aiConfig,
      compactionConfig: this.config.compactionConfig,
    });
  }

  async resumeThread(threadId: ThreadId): Promise<boolean> {
    return resumeThreadSession({
      ...this.threadSessionDeps,
      isRunning: this.turnState.isRunning,
      activeThread: this.turnState.activeThread,
      threadId,
    });
  }

  async sweepIdleThread(now = Date.now()): Promise<boolean> {
    return sweepIdleThreadSession({
      ...this.threadSessionDeps,
      now,
      isRunning: this.turnState.isRunning,
      activeThread: this.turnState.activeThread,
    });
  }

  async runTurn(
    input: AgentTurnInput,
    callbacks: AgentTurnCallbacks
  ): Promise<Turn> {
    return runTurnFlow({
      turnInput: input, callbacks, turnState: this.turnState,
      sessionStore: this.sessionStore, threadStateManager: this.threadStateManager,
      setAutoDrainInputQueue: (enabled) => { this.autoDrainInputQueue = enabled; },
      shouldDrainInputQueue: () => this.autoDrainInputQueue && this.inputQueue.size() > 0,
      scheduleInputQueueDrain: () => this.scheduleInputQueueDrain(),
      startThread: () => this.startThread(),
      clearIdleUnloadTimer: () => this.clearIdleUnloadTimer(),
      publishThreadStatus: () => this.publishThreadStatus(),
      persistThreadRuntime: (threadId) => this.persistThreadRuntime(threadId),
      bindCallbacksToThread: (sourceCallbacks, threadId, clientId) =>
        this.bindCallbacksToThread(sourceCallbacks, threadId, clientId),
      getAllTurnItems: () => this.getAllTurnItems(),
      compactionConfig: this.config.compactionConfig,
      consumePendingCompactionReason: () => this.consumePendingCompactionReason(),
      performAutoCompaction: (thread, reason, targetCallbacks) =>
        this.performAutoCompaction(thread, reason, targetCallbacks),
      persistThreadSnapshot: (thread) => this.persistThreadSnapshot(thread),
      runAgentLoop: (turn, targetCallbacks, turnInput, resumeContext) =>
        this.runAgentLoop(turn, targetCallbacks, turnInput, resumeContext),
      scheduleTurnMemoryExtraction: (thread, turn) =>
        this.scheduleTurnMemoryExtraction(thread, turn),
      scheduleIdleThreadUnload: () => this.scheduleIdleThreadUnload(),
    });
  }

  async interrupt(requestId: ConnectionRequestId = `interrupt-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`): Promise<void> {
    await interruptCurrentTurn({
      requestId, pendingInterruptQueue: this.pendingInterruptQueue, inputQueue: this.inputQueue, turnState: this.turnState,
      disableAutoDrain: () => {
        this.autoDrainInputQueue = false;
      },
    });
  }

  private scheduleInputQueueDrain(): void {
    scheduleQueuedTurnsDrain({
      autoDrainInputQueue: this.autoDrainInputQueue,
      isDrainingInputQueue: this.isDrainingInputQueue,
      isRunning: this.turnState.isRunning,
      setDraining: (isDraining) => { this.isDrainingInputQueue = isDraining; },
      drain: () => this.drainInputQueue(),
    });
  }

  private async drainInputQueue(): Promise<void> {
    await drainQueuedTurnsAndReschedule({
      inputQueue: this.inputQueue,
      isRunning: () => this.turnState.isRunning,
      autoDrainInputQueue: () => this.autoDrainInputQueue,
      runTurn: (input, callbacks) => this.runTurn(input, callbacks),
      setDraining: (isDraining) => { this.isDrainingInputQueue = isDraining; },
      scheduleDrain: () => this.scheduleInputQueueDrain(),
      onTurnError: (error) => agentLoopLogger.warn("执行排队输入失败", error instanceof Error ? { message: error.message, stack: error.stack } : { error: String(error) }),
    });
  }

  private async persistThreadSnapshot(thread: Thread): Promise<void> {
    await persistThreadSnapshotHelper({
      stateRuntimeStore: this.stateRuntimeStore,
      thread,
    });
  }

  private async persistThreadRuntime(threadId: ThreadId): Promise<void> {
    await persistThreadRuntimeHelper({
      stateRuntimeStore: this.stateRuntimeStore,
      snapshot: this.threadStateManager.getSnapshot(),
      threadId,
    });
  }

  private scheduleTurnMemoryExtraction(thread: Thread, turn: Turn): void {
    scheduleTurnMemoryExtractionHelper({
      aiClient: this.aiClient,
      memoryStore: this.config.memoryStore,
      thread,
      turn,
    });
  }

  private publishThreadStatus(): void {
    this.threadWatchManager.publish(this.threadStateManager.getSnapshot());
  }

  private scheduleIdleThreadUnload(): void {
    this.idleUnloadTimer = scheduleIdleThreadUnloadTimer({
      currentTimer: this.idleUnloadTimer,
      isRunning: this.turnState.isRunning,
      hasActiveThread: Boolean(this.turnState.activeThread),
      getStatus: () => this.threadStateManager.getSnapshot(),
      sweepIdleThread: () => this.sweepIdleThread(),
      scheduleAgain: () => this.scheduleIdleThreadUnload(),
    });
  }

  private clearIdleUnloadTimer(): void {
    this.idleUnloadTimer = clearIdleThreadUnloadTimer(this.idleUnloadTimer);
  }

  private async runAgentLoop(
    turn: Turn,
    callbacks: AgentTurnCallbacks,
    input: AgentTurnInput,
    resumeContext?: string
  ): Promise<void> {
    await runAgentLoopWithDeps({
      turn,
      callbacks,
      turnInput: input,
      resumeContext,
      aiClient: this.aiClient,
      aiConfig: this.config.aiConfig,
      configuredReasoningMode: this.config.reasoningMode,
      baseSystemPrompt: this.config.systemPrompt,
      folderId: this.turnState.activeThread?.metadata.folderId,
      stateRuntimeStore: this.stateRuntimeStore,
      toolExecutors: this.config.toolExecutors,
      permissionMode: this.config.permissionMode,
      requestToolApproval: this.config.requestToolApproval,
      samplingRetryConfig: this.config.aiRequestRetryConfig?.sampling,
      signal: this.turnState.abortController?.signal,
      appendTurnItem: (threadId, turnId, item) =>
        this.sessionStore.appendTurnItem(threadId, turnId, item),
      appendToolExecutionLog: (record) =>
        this.stateRuntimeStore?.appendToolExecutionLog(record) ?? Promise.resolve(),
      getTurnItemGroups: () => this.getTurnItemGroups(),
      getActiveThread: () => this.turnState.activeThread,
      getSessionCompactionConfig: () => this.getSessionCompactionConfig(),
      runMidTurnCompaction: () => this.performMidTurnCompaction(turn, callbacks),
      emitContextUsage: (targetCallbacks, requestContext) =>
        this.emitContextUsage(targetCallbacks, requestContext),
      throwIfAborted: () => this.throwIfAborted(),
    });
  }

  private throwIfAborted(): void {
    if (!this.turnState.abortController?.signal.aborted) return;
    const error = new Error("aborted");
    error.name = "AbortError";
    throw error;
  }

  private getSessionCompactionConfig(thread = this.turnState.activeThread): CompactionConfig {
    const globalConfig = this.config.compactionConfig ?? DEFAULT_COMPACTION_CONFIG;
    const contextWindowSize = thread?.metadata.contextWindowSize
      || globalConfig.contextWindowSize
      || 128_000;
    return buildSessionCompactionConfig(globalConfig, contextWindowSize);
  }

  private async generateCompactionSummary(prompt: string): Promise<string> {
    return generateCompactionSummaryWithRetry({
      provider: this.compactionProvider,
      historyPrompt: prompt,
      compactionConfig: this.config.compactionConfig,
      retryOverride: this.config.aiRequestRetryConfig?.compact,
      signal: this.turnState.abortController?.signal,
    });
  }

  private async performAutoCompaction(
    thread: Thread,
    reason: CompactionReason,
    callbacks: AgentTurnCallbacks
  ): Promise<void> {
    await runAutoCompaction({
      thread,
      reason,
      callbacks,
      deps: this.createCompactionRunnerDeps(),
    });
  }

  private async performMidTurnCompaction(
    turn: Turn,
    callbacks: AgentTurnCallbacks
  ): Promise<void> {
    await runMidTurnCompaction({
      turn,
      callbacks,
      deps: this.createCompactionRunnerDeps(),
    });
  }

  private createCompactionRunnerDeps() {
    return createCompactionRunnerDeps({
      sessionStore: this.sessionStore,
      getAllTurnItems: () => this.getAllTurnItems(),
      generateCompactionSummary: (prompt: string) => this.generateCompactionSummary(prompt),
      getSessionCompactionConfig: () => this.getSessionCompactionConfig(),
      archiveRolloutAfterBytes: this.config.compactionConfig?.archiveRolloutAfterBytes,
      setCompactedHistory: (history: TurnItem[]) => { this.turnState.compactedHistory = history; },
      getActiveThread: () => this.turnState.activeThread,
      compactionConfig: this.config.compactionConfig,
    });
  }

  private getAllTurnItems(): TurnItem[] {
    return collectPromptTurnItems({
      activeThread: this.turnState.activeThread,
      activeTurn: this.turnState.activeTurn,
      compactedHistory: this.turnState.compactedHistory,
    });
  }

  private getTurnItemGroups(): TurnItem[][] {
    return collectPromptTurnItemGroups({
      activeThread: this.turnState.activeThread,
      activeTurn: this.turnState.activeTurn,
      compactedHistory: this.turnState.compactedHistory,
    });
  }

  private emitContextUsage(
    callbacks: AgentTurnCallbacks,
    requestContext?: { systemPrompt?: string; tools?: ToolDefinition[] }
  ): void {
    callbacks.onEvent(buildContextUsageEvent({
      groups: this.getTurnItemGroups(),
      activeThread: this.turnState.activeThread,
      compactionConfig: this.config.compactionConfig,
      systemPrompt: requestContext?.systemPrompt,
      tools: requestContext?.tools ?? getToolDefs(this.config.toolExecutors),
    }));
  }

  updateCompactionConfig(config: CompactionConfig): void {
    const result = applyCompactionConfigUpdate({
      currentConfig: this.config,
      nextConfig: config,
      aiClient: this.aiClient,
      activeThread: this.turnState.activeThread,
      usesCustomCompactionProvider: this.usesCustomCompactionProvider,
    });
    if (result.compactionProvider) this.compactionProvider = result.compactionProvider;
    this.markPendingCompactionReason(result.pendingReason);
  }

  async resumeFromInterruption(
    userMessage: string,
    callbacks: AgentTurnCallbacks,
    attachments?: import("../../shared/types").FileAttachment[]
  ): Promise<Turn> {
    const allItems = this.getAllTurnItems();
    const resumeContext = buildResumeContext(allItems);

    return this.runTurn(
      {
        content: userMessage,
        attachments,
        isResume: true,
        resumeContext,
      },
      callbacks
    );
  }
}
