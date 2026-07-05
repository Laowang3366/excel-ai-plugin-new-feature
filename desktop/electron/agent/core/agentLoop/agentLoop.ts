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
  type ToolExecutor,
  type ToolDefinition,
  type CompactionConfig,
  type CompactionReason,
  type CompactProgressItem,
  type ThreadRuntimeSnapshot,
  DEFAULT_COMPACTION_CONFIG,
  mergeTokenUsage,
} from "../../shared/types";

import {
  type AIClientConfig,
  type ReasoningMode,
  createAIClient,
} from "../../providers/aiClient";

import {
  buildResumeContext,
  estimateRequestTokens,
  shouldCompact,
} from "../../memory/compaction";
import { turnItemGroupsToChatMessages } from "../../shared/messageBuilder";

import { SessionStore } from "../../memory/sessionStore";
import type { StateRuntimeStore } from "../../memory/stateRuntimeStore";
import type { LongTermMemoryStore } from "../../memory/longTerm/memoryStore";

// 子模块导入
import { collectStreamEvents, type ToolCallInfo } from "./streamCollector";
import { processToolCalls, getToolDefinitions as getToolDefs, type ToolApprovalConfig } from "./toolExecutor";
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
import { completeTurn, createTurn, createUserMessageItem } from "./turnRunner";
import { PendingInterruptQueue, type ConnectionRequestId } from "./pendingInterruptQueue";
import { ThreadWatchManager } from "./threadWatchManager";
import { InputQueue } from "./inputQueue";
import { isModelCompHashCompatible, resolveModelCompHash } from "./modelCompHash";
import {
  buildContextUsageEvent,
  collectPromptTurnItemGroups,
  collectPromptTurnItems,
} from "./contextUsage";
import { emitStreamResultItems as emitCollectedStreamResultItems } from "./streamResultItems";
import {
  archiveRolloutIfConfigured as archiveRolloutIfConfiguredHelper,
  completeCompactionProgress as completeCompactionProgressHelper,
  failCompactionProgress as failCompactionProgressHelper,
  startCompactionProgress as startCompactionProgressHelper,
} from "./compactionProgress";
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
import { buildRoundStreamParams } from "./roundStreamParams";
import {
  resetThreadSession,
  resumeThreadSession,
  startThreadSession,
  sweepIdleThreadSession,
} from "./threadSession";
import {
  DEFAULT_COMPACT_RETRY_CONFIG,
  runAIRequestWithRetry,
  type AIRequestPhase,
  type AIRequestRetryConfig,
} from "./aiRequestRetry";

// ============================================================
// Agent Loop 配置
// ============================================================

export interface AgentLoopConfig {
  /** AI 客户端配置 */
  aiConfig: AIClientConfig;
  /** 系统提示词 */
  systemPrompt?: string;
  /** 压缩配置 */
  compactionConfig?: CompactionConfig;
  /** 压缩摘要生成器；默认使用当前 AI 客户端本地生成 */
  compactionProvider?: CompactionProvider;
  /** 工具执行器映射 */
  toolExecutors?: Map<string, ToolExecutor>;
  /** 权限模式 */
  permissionMode?: "normal" | "auto_approve_safe" | "confirm_all";
  /** 推理力度（覆盖 aiConfig 中的 reasoningMode */
  reasoningMode?: ReasoningMode;
  /** 空闲多久后卸载内存中的 activeThread，默认 30 分钟；<=0 表示禁用 */
  threadIdleUnloadMs?: number;
  /** 请求工具审批的回调（由主进程提供，向渲染进程发送审批请求并等待响应） */
  requestToolApproval?: (params: {
    toolCallId: string;
    toolName: string;
    arguments: Record<string, unknown>;
    riskLevel: "safe" | "moderate" | "dangerous";
    description?: string;
  }) => Promise<{ approved: boolean; alwaysAllow?: boolean }>;
  /** 模型请求重试配置。sampling 用于普通回复，compact 用于上下文压缩。 */
  aiRequestRetryConfig?: Partial<Record<AIRequestPhase, AIRequestRetryConfig>>;
  /** Turn 完成后自动抽取并写入长期记忆。 */
  memoryStore?: LongTermMemoryStore;
}
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
  private idleUnloadTimer: ReturnType<typeof setTimeout> | null = null;
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
    if (!this.autoDrainInputQueue) {
      throw new Error("Agent 正在中断中，请等待停止完成后再发送新请求");
    }
    const queueSize = this.inputQueue.enqueue({ input, callbacks });
    return { queued: true, queueSize };
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
    const previous = this.config.aiConfig;
    this.config.aiConfig = config;
    this.aiClient = createAIClient(config);
    if (!this.usesCustomCompactionProvider) {
      this.compactionProvider = createCompactionProvider(this.aiClient, this.config.compactionConfig);
    }
    if (this.turnState.activeThread) {
      this.turnState.activeThread.metadata.modelProvider = config.provider;
      this.turnState.activeThread.metadata.model = config.model;
      this.turnState.activeThread.metadata.contextWindowSize = config.contextWindowSize ?? this.turnState.activeThread.metadata.contextWindowSize;
      this.turnState.activeThread.metadata.compHash = resolveModelCompHash(config);
    }
    if (this.turnState.activeThread && !isModelCompHashCompatible(previous, config)) {
      this.markPendingCompactionReason("model_changed");
    }
    if (this.turnState.activeThread && previous.contextWindowSize !== config.contextWindowSize) {
      this.markPendingCompactionReason("context_window_changed");
    }
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
    if (this.turnState.isRunning) {
      throw new Error("Agent 正在运行中，请等待当前 Turn 完成或中断");
    }

    this.turnState.isRunning = true;
    this.autoDrainInputQueue = true;
    this.turnState.abortController = new AbortController();

    // 创建 turn 完成 Promise，供 interrupt() 等待清理完毕
    this.turnState.turnCompletionPromise = new Promise((resolve) => {
      this.turnState.resolveTurnCompletion = resolve;
    });

    let turnCallbacks = callbacks;

    try {
      // 确保有活跃线程
      if (!this.turnState.activeThread) {
        await this.startThread();
      }
      const thread = this.turnState.activeThread!;
      this.clearIdleUnloadTimer();
      this.threadStateManager.markRunning(thread.metadata.threadId);
      this.publishThreadStatus();
      await this.persistThreadRuntime(thread.metadata.threadId);
      turnCallbacks = this.bindCallbacksToThread(callbacks, thread.metadata.threadId, input.clientId);

      // 检查是否需要压缩（参考 Codex 的 pre-turn compaction）
      // 使用会话自己的 contextWindowSize 构建压缩配置，确保会话间隔离
      const allItems = this.getAllTurnItems();
      const globalConfig = this.config.compactionConfig ?? DEFAULT_COMPACTION_CONFIG;
      const sessionContextWindowSize = thread.metadata.contextWindowSize
        || globalConfig.contextWindowSize
        || 128_000;
      const sessionCompactionConfig = buildSessionCompactionConfig(globalConfig, sessionContextWindowSize);
      const pendingReason = this.consumePendingCompactionReason(allItems, sessionCompactionConfig);
      if (pendingReason) {
        await this.performAutoCompaction(thread, pendingReason, turnCallbacks);
      } else if (shouldCompact(allItems, sessionCompactionConfig)) {
        await this.performAutoCompaction(thread, "auto_pre_turn", turnCallbacks);
      }

      // 创建新 Turn
      const turn = createTurn(thread.metadata.threadId);
      this.turnState.activeTurn = turn;
      thread.metadata.activeTurnId = turn.turnId;
      thread.metadata.lastTurnStatus = "in_progress";
      await this.persistThreadSnapshot(thread);

      // 发送 Turn 开始事件
      turnCallbacks.onEvent({ type: "turn_started", turnId: turn.turnId });

      // 添加用户消息并通过事件发出（参考 Codex：消息只从 agent 事件产出）
      const userItem = createUserMessageItem(input);
      turn.items.push(userItem);
      await this.sessionStore.appendTurnItem(thread.metadata.threadId, turn.turnId, userItem);
      // 通过事件通知前端——前端不再自行创建用户消息
      turnCallbacks.onEvent({ type: "item_started", item: userItem });
      turnCallbacks.onEvent({ type: "item_completed", item: userItem });

      // 更新会话预览
      if (!thread.metadata.preview) {
        thread.metadata.preview = input.content.slice(0, 100);
      }

      // 运行 Agent 循环，传入恢复上下文（仅系统内部使用，不显示给用户）
      const resumeContext = input.isResume ? input.resumeContext : undefined;
      await this.runAgentLoop(turn, turnCallbacks, input, resumeContext);

      // 完成 Turn
      completeTurn(turn);
      if (turn.tokenUsage) {
        await this.sessionStore.appendTurnUsage(thread.metadata.threadId, turn.turnId, turn.tokenUsage);
      }
      thread.turns.push(turn);
      thread.metadata.updatedAt = Date.now();
      thread.metadata.lastTurnStatus = turn.status;
      thread.metadata.activeTurnId = undefined;
      await this.persistThreadSnapshot(thread);

      turnCallbacks.onEvent({
        type: "turn_completed",
        turnId: turn.turnId,
        usage: turn.tokenUsage,
      });
      this.scheduleTurnMemoryExtraction(thread, turn);

      return turn;
    } catch (err: any) {
      if (this.turnState.activeTurn) {
        this.turnState.activeTurn.status = err.name === "AbortError" ? "interrupted" : "failed";
        this.turnState.activeTurn.error = err.message;
        this.turnState.activeTurn.completedAt = Date.now();

        if (this.turnState.activeTurn.status === "interrupted") {
          turnCallbacks.onEvent({ type: "turn_interrupted", turnId: this.turnState.activeTurn.turnId });
        } else {
          turnCallbacks.onEvent({
            type: "turn_failed",
            turnId: this.turnState.activeTurn.turnId,
            error: err.message,
          });
        }

        this.turnState.activeThread?.turns.push(this.turnState.activeTurn);
        if (this.turnState.activeThread) {
          this.turnState.activeThread.metadata.updatedAt = this.turnState.activeTurn.completedAt ?? Date.now();
          this.turnState.activeThread.metadata.lastTurnStatus = this.turnState.activeTurn.status;
          this.turnState.activeThread.metadata.activeTurnId = undefined;
          await this.persistThreadSnapshot(this.turnState.activeThread);
        }
      }
      throw err;
    } finally {
      this.turnState.isRunning = false;
      this.turnState.abortController = null;
      if (this.turnState.activeThread) {
        this.threadStateManager.markIdle(this.turnState.activeThread.metadata.threadId);
        this.publishThreadStatus();
        this.scheduleIdleThreadUnload();
        await this.persistThreadRuntime(this.turnState.activeThread.metadata.threadId);
      }
      if (this.autoDrainInputQueue && this.inputQueue.size() > 0) {
        this.scheduleInputQueueDrain();
      }
      // 通知 interrupt() 等待者：Turn 清理已完成
      this.turnState.resolveTurnCompletion?.();
      this.turnState.turnCompletionPromise = null;
      this.turnState.resolveTurnCompletion = null;
    }
  }

  /** 中断当前 Turn，等待清理完成后返回 */
  async interrupt(requestId: ConnectionRequestId = `interrupt-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`): Promise<void> {
    this.pendingInterruptQueue.push(requestId);
    this.autoDrainInputQueue = false;
    this.inputQueue.clear();
    this.turnState.abortController?.abort();
    // 等待 runTurn() 的 catch/finally 完全执行完毕，
    // 确保 isRunning=false、activeThread/activeTurn 稳定后才返回。
    try {
      if (this.turnState.turnCompletionPromise) {
        await this.turnState.turnCompletionPromise;
      }
    } finally {
      this.inputQueue.clear();
      this.pendingInterruptQueue.drain();
    }
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
      while (!this.turnState.isRunning) {
        const next = this.inputQueue.dequeue();
        if (!next) return;
        try {
          await this.runTurn(next.input, next.callbacks);
        } catch (error) {
          console.warn("执行排队输入失败:", error);
        }
      }
    } finally {
      this.isDrainingInputQueue = false;
      if (this.autoDrainInputQueue && !this.turnState.isRunning && this.inputQueue.size() > 0) {
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
    this.clearIdleUnloadTimer();
    if (this.turnState.isRunning || !this.turnState.activeThread) return;

    const status = this.threadStateManager.getSnapshot();
    if (status.idleUnloadMs <= 0 || status.lastActiveAt === undefined) return;

    const delay = Math.max(0, status.idleUnloadMs - (Date.now() - status.lastActiveAt));
    this.idleUnloadTimer = setTimeout(() => {
      void this.sweepIdleThread().catch(() => {
        this.scheduleIdleThreadUnload();
      });
    }, delay);
    (this.idleUnloadTimer as { unref?: () => void }).unref?.();
  }

  private clearIdleUnloadTimer(): void {
    if (!this.idleUnloadTimer) return;
    clearTimeout(this.idleUnloadTimer);
    this.idleUnloadTimer = null;
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
    let round = 0;

    while (true) {
      this.throwIfAborted();
      round++;

      const { streamParams, effectiveSystemPrompt, toolDefs } = await buildRoundStreamParams({
        turnItemGroups: this.getTurnItemGroups(),
        turnInput: input,
        aiConfig: this.config.aiConfig,
        configuredReasoningMode: this.config.reasoningMode,
        baseSystemPrompt: this.config.systemPrompt,
        folderId: this.turnState.activeThread?.metadata.folderId,
        stateRuntimeStore: this.stateRuntimeStore,
        toolExecutors: this.config.toolExecutors,
        signal: this.turnState.abortController?.signal,
        round,
        resumeContext,
      });

      // 3. 收集响应
      let streamResult;
      try {
        streamResult = await runAIRequestWithRetry({
          phase: "sampling",
          config: this.config.aiRequestRetryConfig?.sampling,
          signal: this.turnState.abortController?.signal,
          operation: () => collectStreamEvents(
            this.aiClient.streamChat(streamParams),
            callbacks,
            round
          ),
        });
      } catch (err: any) {
        if (err.name === "AbortError") {
          // 中断处理已移到 streamCollector
          throw err;
        }
        throw err;
      }

      // 处理流式错误事件
      const errorItem = (streamResult as any).errorItem as TurnItem | undefined;
      if (errorItem) {
        turn.items.push(errorItem);
        await this.sessionStore.appendTurnItem(turn.threadId, turn.turnId, errorItem);
        callbacks.onEvent({ type: "item_started", item: errorItem });
        callbacks.onEvent({ type: "item_completed", item: errorItem });
        callbacks.onEvent({ type: "error", message: errorItem.type === "error" ? errorItem.message : "Unknown error" });
        return;
      }

      // 处理中断（AbortError 在 collectStreamEvents 外部抛出）
      // 这里不会到这里——AbortError 会在上方 throw

      // 4. 流结束，按 API 真实事件顺序发出 items
      this.throwIfAborted();
      await emitCollectedStreamResultItems({
        streamResult,
        turn,
        callbacks,
        appendTurnItem: (threadId, turnId, item) =>
          this.sessionStore.appendTurnItem(threadId, turnId, item),
      });
      this.throwIfAborted();

      // 5. 处理工具调用
      if (streamResult.toolCalls.length > 0) {
        const approvalConfig: ToolApprovalConfig = {
          permissionMode: this.config.permissionMode || "normal",
          requestToolApproval: this.config.requestToolApproval,
        };

        await processToolCalls(
          streamResult.toolCalls,
          streamResult.pendingToolCallItems,
          turn,
          this.config.toolExecutors!,
          approvalConfig,
          callbacks,
          async (threadId, turnId, item) => {
            await this.sessionStore.appendTurnItem(threadId, turnId, item);
          },
          async (record) => {
            await this.stateRuntimeStore?.appendToolExecutionLog(record);
          }
        );
        this.throwIfAborted();

        // 检查是否需要压缩（mid-turn compaction）
        const compactionConfig = this.getSessionCompactionConfig();
        const allTokens = estimateRequestTokens({
          messages: turnItemGroupsToChatMessages(this.getTurnItemGroups()),
          systemPrompt: effectiveSystemPrompt,
          tools: toolDefs,
        });
        const midTurnRatio = compactionConfig.midTurnThresholdRatio ?? 0.9;
        if (compactionConfig.enabled && allTokens > compactionConfig.autoCompactTokenThreshold * midTurnRatio) {
          await this.performMidTurnCompaction(turn, callbacks);
        }
        this.throwIfAborted();

        // 继续循环
        continue;
      }

      // 7. 没有工具调用 → Turn 结束
      if (streamResult.usage) {
        turn.tokenUsage = streamResult.usage;
        if (this.turnState.activeThread) {
          this.turnState.activeThread.metadata.totalTokenUsage = this.turnState.activeThread.metadata.totalTokenUsage
            ? mergeTokenUsage(this.turnState.activeThread.metadata.totalTokenUsage, streamResult.usage)
            : streamResult.usage;
        }
      }
      // Turn 结束时始终发送上下文使用情况
      this.emitContextUsage(callbacks, {
        systemPrompt: effectiveSystemPrompt,
        tools: toolDefs,
      });
      break;
    }
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
    const config = this.config.compactionConfig ?? DEFAULT_COMPACTION_CONFIG;
    const retryCount = Math.max(
      0,
      Math.floor(config.summaryRetryCount ?? DEFAULT_COMPACTION_CONFIG.summaryRetryCount ?? 0)
    );
    const compactRetryConfig: AIRequestRetryConfig = {
      maxRetries: retryCount,
      baseDelayMs: config.summaryRetryBaseDelayMs ?? DEFAULT_COMPACT_RETRY_CONFIG.baseDelayMs,
      maxDelayMs: config.summaryRetryMaxDelayMs ?? DEFAULT_COMPACT_RETRY_CONFIG.maxDelayMs,
      backoffFactor: config.summaryRetryBackoffFactor ?? DEFAULT_COMPACT_RETRY_CONFIG.backoffFactor,
      ...this.config.aiRequestRetryConfig?.compact,
    };

    return runAIRequestWithRetry({
      phase: "compact",
      config: compactRetryConfig,
      signal: this.turnState.abortController?.signal,
      operation: () => this.compactionProvider.generateSummary({
        historyPrompt: prompt,
        config,
      }),
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
    return {
      sessionStore: this.sessionStore,
      getAllTurnItems: () => this.getAllTurnItems(),
      generateCompactionSummary: (prompt: string) => this.generateCompactionSummary(prompt),
      startCompactionProgress: (
        threadId: ThreadId,
        reason: CompactionReason,
        items: TurnItem[],
        callbacks: AgentTurnCallbacks
      ) => startCompactionProgressHelper({
        sessionStore: this.sessionStore,
        threadId,
        reason,
        items,
        callbacks,
        compactionConfig: this.getSessionCompactionConfig(),
      }),
      completeCompactionProgress: (
        progress: CompactProgressItem,
        tokensBefore: number,
        tokensAfter: number,
        summary: string,
        callbacks: AgentTurnCallbacks
      ) => completeCompactionProgressHelper({
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
      ) => failCompactionProgressHelper({
        sessionStore: this.sessionStore,
        threadId,
        progress,
        items,
        error,
        callbacks,
      }),
      archiveRolloutIfConfigured: (threadId: ThreadId) => archiveRolloutIfConfiguredHelper({
        sessionStore: this.sessionStore,
        threadId,
        threshold: this.config.compactionConfig?.archiveRolloutAfterBytes,
      }),
      setCompactedHistory: (history: TurnItem[]) => {
        this.turnState.compactedHistory = history;
      },
      getActiveThread: () => this.turnState.activeThread,
      compactionConfig: this.config.compactionConfig,
    };
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
    const previousWindow = this.config.compactionConfig?.contextWindowSize;
    this.config.compactionConfig = config;
    if (!this.usesCustomCompactionProvider) {
      this.compactionProvider = createCompactionProvider(this.aiClient, config);
    }
    if (this.turnState.activeThread && previousWindow !== config.contextWindowSize) {
      this.markPendingCompactionReason("context_window_changed");
    }
  }

  private markPendingCompactionReason(reason: CompactionReason): void {
    if (this.pendingCompactionReason === "model_changed" && reason === "context_window_changed") {
      return;
    }
    this.pendingCompactionReason = reason;
  }

  private consumePendingCompactionReason(
    items: TurnItem[],
    config: CompactionConfig
  ): CompactionReason | null {
    const reason = this.pendingCompactionReason;
    this.pendingCompactionReason = null;
    if (!reason || !config.enabled || items.length === 0) return null;
    return reason;
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
