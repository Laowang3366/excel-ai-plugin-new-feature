/**
 * Agent 循环 — 参考 Codex 的 turn.rs 和 codex_thread.rs 重写
 *
 * 核心流程（参考 Codex run_turn）：
 *   1. 用户发送消息 → 创建新 Turn
 *   2. 构建 AI 请求（包含历史上下文）
 *   3. 流式接收 AI 响应
 *      - 文本 → 直接输出
 *      - 推理 → 完整展示（不隐藏）
 *      - 工具调用 → 执行工具 → 将结果发回 AI → 继续
 *   4. AI 不再请求工具 → Turn 完成
 *
 * 关键设计（参考 Codex 事件驱动模型）：
 *   - 消息列表是 Agent 事件的唯一投影，前端不自行创建消息
 *   - 所有 TurnItem 通过 item_started/item_completed 事件发出
 *   - 设置变更只更新 AI 客户端，不销毁线程
 *   - 压缩历史在 thread 级别维护，不创建虚拟 Turn
 */

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
  type ThreadRuntimeSnapshot,
  DEFAULT_COMPACTION_CONFIG,
} from "../../shared/types";

import {
  type AIClientConfig,
  createAIClient,
} from "../../providers/aiClient";

import {
  buildResumeContext,
} from "../../memory/compaction";

import { SessionStore } from "../../memory/sessionStore";
import type { StateRuntimeStore } from "../../memory/stateRuntimeStore";

// 子模块导入
import { getToolDefinitions as getToolDefs } from "./toolExecutor";
import { buildSessionCompactionConfig } from "./sessionCompactionConfig";
import { TurnState } from "./turnState";
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
import {
  buildContextUsageEvent,
  collectPromptTurnItemGroups,
  collectPromptTurnItems,
} from "./contextUsage";
import {
  attachRolloutEventSink as attachRolloutEventSinkHelper,
  bindCallbacksToThread as bindCallbacksToThreadHelper,
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
  drainQueuedTurns,
  enqueueQueuedTurn,
  interruptCurrentTurn,
  shouldRescheduleQueueDrain,
} from "./queuedTurns";
import {
  applyAIConfigUpdate,
  applyCompactionConfigUpdate,
  mergePendingCompactionReason,
} from "./configUpdates";
import { generateCompactionSummary as generateCompactionSummaryWithRetry } from "./compactionSummary";
import { runAgentLoopRounds } from "./agentLoopRunner";
import { createCompactionRunnerDeps } from "./compactionRunnerDeps";
import {
  clearIdleThreadUnloadTimer,
  scheduleIdleThreadUnload as scheduleIdleThreadUnloadTimer,
  type IdleThreadUnloadTimer,
} from "./idleThreadUnload";
import { runTurnFlow } from "./turnFlow";
import type { AgentLoopConfig } from "./agentLoopConfig";

export type { AgentLoopConfig } from "./agentLoopConfig";

// ============================================================
// Agent Loop — 核心循环
// ============================================================

export class AgentLoop {
  private config: AgentLoopConfig;
  private aiClient: ReturnType<typeof createAIClient>;
  private sessionStore: SessionStore;
  private stateRuntimeStore: StateRuntimeStore | undefined;
  private readonly turnState = new TurnState();
  private readonly threadStateManager: ThreadStateManager;
  private readonly pendingInterruptQueue = new PendingInterruptQueue();
  private readonly threadWatchManager: ThreadWatchManager;
  private readonly inputQueue = new InputQueue();
  private compactionProvider: CompactionProvider;
  private readonly usesCustomCompactionProvider: boolean;
  private pendingCompactionReason: CompactionReason | null = null;
  private idleUnloadTimer: IdleThreadUnloadTimer | null = null;
  private isDrainingInputQueue = false;
  private autoDrainInputQueue = true;

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

  // ----------------------------------------------------------
  // 公共 API
  // ----------------------------------------------------------

  /** 迁移后更新 SessionStore 引用（数据迁移场景） */
  updateSessionStore(store: SessionStore): void {
    this.sessionStore = store;
    this.attachRolloutEventSink();
  }

  /** 数据迁移后更新 SQLite 运行态存储引用。 */
  updateStateRuntimeStore(store: StateRuntimeStore): void {
    this.stateRuntimeStore = store;
    this.config.memoryStore?.updateRuntime(store);
    this.attachRolloutEventSink();
  }

  /** 获取当前活跃的线程 */
  getThread(): Thread | null {
    return this.turnState.activeThread;
  }

  /** 获取当前线程运行态快照，用于状态观察和诊断。 */
  getThreadRuntimeStatus(): ThreadRuntimeSnapshot {
    return this.threadStateManager.getSnapshot();
  }

  /** 是否正在运行 */
  getIsRunning(): boolean {
    return this.turnState.isRunning;
  }

  /** 当前排队中的中断请求 ID，用于诊断多来源中断。 */
  getPendingInterruptRequestIds(): ConnectionRequestId[] {
    return this.pendingInterruptQueue.pendingIds();
  }

  /** 运行中用户补充输入数量。 */
  getQueuedInputCount(): number {
    return this.inputQueue.size();
  }

  /** 当前 turn 运行中收到的新输入先入队，待当前 turn 结束后自动继续。 */
  enqueueTurn(
    input: AgentTurnInput,
    callbacks: AgentTurnCallbacks
  ): { queued: true; queueSize: number } {
    return enqueueQueuedTurn({ autoDrainInputQueue: this.autoDrainInputQueue, inputQueue: this.inputQueue, turnInput: input, callbacks });
  }

  watchThreadStatus(
    threadId: ThreadId,
    connectionId: string,
    listener: (status: ThreadRuntimeSnapshot) => void
  ) {
    return this.threadWatchManager.watch(threadId, connectionId, listener);
  }

  private attachRolloutEventSink(): void {
    attachRolloutEventSinkHelper({
      sessionStore: this.sessionStore,
      stateRuntimeStore: this.stateRuntimeStore,
    });
  }

  /**
   * 更新 AI 客户端配置（参考 Codex CodexThreadSettingsOverrides）
   *
   * 设置变更时只更新 AI 客户端，不销毁线程。
   * 这样切换 provider 后对话历史不会丢失。
   */
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

  /** 更新权限模式（热更新，不销毁线程） */
  updatePermissionMode(mode: "normal" | "auto_approve_safe" | "confirm_all"): void {
    this.config.permissionMode = mode;
  }

  private bindCallbacksToThread(
    callbacks: AgentTurnCallbacks,
    threadId: ThreadId,
    clientId?: string
  ): AgentTurnCallbacks {
    return bindCallbacksToThreadHelper({ callbacks, threadId, clientId });
  }

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

  /**
   * 重置线程（用于新建会话）
   *
   * 如果当前有正在运行的 Turn，先中断并等待清理完成。
   * 清除 activeThread，下次 startTurn 时自动创建新线程。
   * 不影响 SessionStore 中的历史数据。
   */
  async resetThread(folderId?: string): Promise<void> {
    await resetThreadSession({
      ...this.threadSessionDeps,
      isRunning: this.turnState.isRunning,
      interrupt: () => this.interrupt(),
      folderId,
    });
  }

  /** 启动新会话 */
  async startThread(): Promise<ThreadId> {
    return startThreadSession({
      ...this.threadSessionDeps,
      aiConfig: this.config.aiConfig,
      compactionConfig: this.config.compactionConfig,
    });
  }

  /** 恢复已有会话 */
  async resumeThread(threadId: ThreadId): Promise<boolean> {
    return resumeThreadSession({
      ...this.threadSessionDeps,
      isRunning: this.turnState.isRunning,
      activeThread: this.turnState.activeThread,
      threadId,
    });
  }

  /** 扫描并卸载空闲线程；返回 true 表示已释放 activeThread。 */
  async sweepIdleThread(now = Date.now()): Promise<boolean> {
    return sweepIdleThreadSession({
      ...this.threadSessionDeps,
      now,
      isRunning: this.turnState.isRunning,
      activeThread: this.turnState.activeThread,
    });
  }

  /** 执行一次 Turn（核心方法） */
  async runTurn(
    input: AgentTurnInput,
    callbacks: AgentTurnCallbacks
  ): Promise<Turn> {
    return runTurnFlow({
      turnInput: input,
      callbacks,
      turnState: this.turnState,
      sessionStore: this.sessionStore,
      threadStateManager: this.threadStateManager,
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
      consumePendingCompactionReason: () => {
        const reason = this.pendingCompactionReason;
        this.pendingCompactionReason = null;
        return reason;
      },
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

  /** 中断当前 Turn，等待清理完成后返回 */
  async interrupt(requestId: ConnectionRequestId = `interrupt-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`): Promise<void> {
    await interruptCurrentTurn({
      requestId, pendingInterruptQueue: this.pendingInterruptQueue, inputQueue: this.inputQueue, turnState: this.turnState,
      disableAutoDrain: () => {
        this.autoDrainInputQueue = false;
      },
    });
  }

  private scheduleInputQueueDrain(): void {
    if (!this.autoDrainInputQueue || this.isDrainingInputQueue || this.turnState.isRunning) return;
    this.isDrainingInputQueue = true;
    queueMicrotask(() => {
      void this.drainInputQueue();
    });
  }

  private async drainInputQueue(): Promise<void> {
    try {
      await drainQueuedTurns({
        inputQueue: this.inputQueue,
        isRunning: () => this.turnState.isRunning,
        runTurn: (input, callbacks) => this.runTurn(input, callbacks),
        onTurnError: (error) => console.warn("执行排队输入失败:", error),
      });
    } finally {
      this.isDrainingInputQueue = false;
      if (shouldRescheduleQueueDrain({ autoDrainInputQueue: this.autoDrainInputQueue, isRunning: this.turnState.isRunning, queueSize: this.inputQueue.size() })) {
        this.scheduleInputQueueDrain();
      }
    }
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

  // ----------------------------------------------------------
  // Agent 循环核心
  // ----------------------------------------------------------

  /**
   * Agent 循环（参考 Codex run_turn）
   *
   * 流程：
   *   while (true) {
   *     1. 构建消息历史
   *     2. 调用 AI
   *     3. 收集响应（文本/推理/工具调用）
   *     4. 如果有工具调用 → 执行 → 添加结果到历史 → continue
   *     5. 如果没有工具调用 → break
   *   }
   *
   * 关键：所有消息通过 item_started/item_completed 事件发出，
   * 前端只从事件获取消息，不自行创建。
   */
  private async runAgentLoop(
    turn: Turn,
    callbacks: AgentTurnCallbacks,
    input: AgentTurnInput,
    resumeContext?: string
  ): Promise<void> {
    await runAgentLoopRounds({
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
      toolExecutors: this.config.toolExecutors!,
      approvalConfig: {
        permissionMode: this.config.permissionMode || "normal",
        requestToolApproval: this.config.requestToolApproval,
      },
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

  /** Pre-turn 自动压缩 */
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

  /** Mid-turn 压缩（参考 Codex 的 mid-turn compaction） */
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

  // ----------------------------------------------------------
  // 历史管理
  // ----------------------------------------------------------

  /**
   * 获取所有 Turn 的条目（参考 Codex ContextManager.for_prompt）
   *
   * 优先使用 compactedHistory（压缩后的历史），
   * 否则从所有 turns 收集。
   * 加上当前活跃 turn 的新条目。
   *
   * invariant: 当 compactedHistory 非空时，thread.turns 只包含
   * 压缩点之后的已完成 turns（或为空），不会与 compactedHistory 重叠。
   */
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

  // ----------------------------------------------------------
  // 上下文使用量（参考 Codex context_window_line / AutoCompactTokenStatus）
  // ----------------------------------------------------------

  /**
   * 发送上下文使用情况事件
   */
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

  /**
   * 更新压缩配置（热更新，不销毁线程）
   */
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

  private markPendingCompactionReason(reason: CompactionReason | null): void {
    this.pendingCompactionReason = mergePendingCompactionReason(this.pendingCompactionReason, reason);
  }

  // ----------------------------------------------------------
  // 中断恢复
  // ----------------------------------------------------------

  /**
   * 从中断处恢复
   */
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
