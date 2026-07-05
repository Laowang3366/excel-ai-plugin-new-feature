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
import { resolveImageAttachments } from "../../attachments/imageAttachmentResolver";

import { SessionStore } from "../../memory/sessionStore";
import type { StateRuntimeStore } from "../../memory/stateRuntimeStore";
import type { LongTermMemoryStore } from "../../memory/longTerm/memoryStore";

// 子模块导入
import { collectStreamEvents, type ToolCallInfo } from "./streamCollector";
import { processToolCalls, getToolDefinitions as getToolDefs, type ToolApprovalConfig } from "./toolExecutor";
import { buildSessionCompactionConfig } from "./sessionCompactionConfig";
import {
  appendRuntimeLongTermMemoryContext,
  buildEffectiveSystemPrompt,
} from "./buildStreamParams";
import { resolveMaxTokens } from "./maxTokens";
import { TurnState } from "./turnState";
import {
  createCompactionProvider,
  type CompactionProvider,
} from "./compactionProvider";
import {
  DEFAULT_THREAD_IDLE_UNLOAD_MS,
  ThreadStateManager,
} from "./threadStateManager";
import { createAgentThread, loadAgentThread } from "./threadLifecycle";
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

  private get activeThread(): Thread | null { return this.turnState.activeThread; }
  private set activeThread(thread: Thread | null) { this.turnState.activeThread = thread; }
  private get activeTurn(): Turn | null { return this.turnState.activeTurn; }
  private set activeTurn(turn: Turn | null) { this.turnState.activeTurn = turn; }
  private get isRunning(): boolean { return this.turnState.isRunning; }
  private set isRunning(value: boolean) { this.turnState.isRunning = value; }
  private get abortController(): AbortController | null { return this.turnState.abortController; }
  private set abortController(value: AbortController | null) { this.turnState.abortController = value; }
  private get _turnCompletionPromise(): Promise<void> | null { return this.turnState.turnCompletionPromise; }
  private set _turnCompletionPromise(value: Promise<void> | null) { this.turnState.turnCompletionPromise = value; }
  private get _resolveTurnCompletion(): (() => void) | null { return this.turnState.resolveTurnCompletion; }
  private set _resolveTurnCompletion(value: (() => void) | null) { this.turnState.resolveTurnCompletion = value; }
  private get compactedHistory(): TurnItem[] | null { return this.turnState.compactedHistory; }
  private set compactedHistory(history: TurnItem[] | null) { this.turnState.compactedHistory = history; }

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
    return this.activeThread;
  }

  /** 获取当前线程运行态快照，用于状态观察和诊断。 */
  getThreadRuntimeStatus(): ThreadRuntimeSnapshot {
    return this.threadStateManager.getSnapshot();
  }

  /** 是否正在运行 */
  getIsRunning(): boolean {
    return this.isRunning;
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
    if (this.activeThread) {
      this.activeThread.metadata.modelProvider = config.provider;
      this.activeThread.metadata.model = config.model;
      this.activeThread.metadata.contextWindowSize = config.contextWindowSize ?? this.activeThread.metadata.contextWindowSize;
      this.activeThread.metadata.compHash = resolveModelCompHash(config);
    }
    if (this.activeThread && !isModelCompHashCompatible(previous, config)) {
      this.markPendingCompactionReason("model_changed");
    }
    if (this.activeThread && previous.contextWindowSize !== config.contextWindowSize) {
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

  /**
   * 重置线程（用于新建会话）
   *
   * 如果当前有正在运行的 Turn，先中断并等待清理完成。
   * 清除 activeThread，下次 startTurn 时自动创建新线程。
   * 不影响 SessionStore 中的历史数据。
   */
  async resetThread(folderId?: string): Promise<void> {
    if (this.isRunning) {
      await this.interrupt();
    }
    this.clearIdleUnloadTimer();
    this.turnState.resetForNextThread(folderId);
    this.threadStateManager.clear();
  }

  /** 启动新会话 */
  async startThread(): Promise<ThreadId> {
    const folderId = this.turnState.consumePendingFolderId();

    const thread = await createAgentThread({
      sessionStore: this.sessionStore,
      aiConfig: this.config.aiConfig,
      compactionConfig: this.config.compactionConfig,
      folderId,
    });
    this.activeThread = thread;
    this.compactedHistory = null;
    this.threadStateManager.markLoaded(thread.metadata.threadId);
    this.publishThreadStatus();
    this.scheduleIdleThreadUnload();
    await this.persistThreadSnapshot(thread);
    await this.persistThreadRuntime(thread.metadata.threadId);
    return thread.metadata.threadId;
  }

  /** 恢复已有会话 */
  async resumeThread(threadId: ThreadId): Promise<boolean> {
    if (this.isRunning) {
      return this.activeThread?.metadata.threadId === threadId;
    }
    const result = await loadAgentThread(this.sessionStore, threadId);
    if (!result) return false;
    this.activeThread = result.thread;
    this.compactedHistory = result.compactedHistory;
    this.threadStateManager.markLoaded(result.thread.metadata.threadId);
    this.publishThreadStatus();
    this.scheduleIdleThreadUnload();
    await this.persistThreadSnapshot(result.thread);
    await this.persistThreadRuntime(result.thread.metadata.threadId);

    return true;
  }

  /** 扫描并卸载空闲线程；返回 true 表示已释放 activeThread。 */
  async sweepIdleThread(now = Date.now()): Promise<boolean> {
    if (this.isRunning || !this.activeThread || !this.threadStateManager.shouldUnload(now)) {
      return false;
    }

    const threadId = this.activeThread.metadata.threadId;
    await this.sessionStore.flushRolloutWrites();
    this.activeThread = null;
    this.activeTurn = null;
    this.compactedHistory = null;
    this.threadStateManager.markUnloaded(now);
    this.publishThreadStatus();
    this.clearIdleUnloadTimer();
    await this.persistThreadRuntime(threadId);
    return true;
  }

  /** 执行一次 Turn（核心方法） */
  async runTurn(
    input: AgentTurnInput,
    callbacks: AgentTurnCallbacks
  ): Promise<Turn> {
    if (this.isRunning) {
      throw new Error("Agent 正在运行中，请等待当前 Turn 完成或中断");
    }

    this.isRunning = true;
    this.autoDrainInputQueue = true;
    this.abortController = new AbortController();

    // 创建 turn 完成 Promise，供 interrupt() 等待清理完毕
    this._turnCompletionPromise = new Promise((resolve) => {
      this._resolveTurnCompletion = resolve;
    });

    let turnCallbacks = callbacks;

    try {
      // 确保有活跃线程
      if (!this.activeThread) {
        await this.startThread();
      }
      const thread = this.activeThread!;
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
      this.activeTurn = turn;
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
      if (this.activeTurn) {
        this.activeTurn.status = err.name === "AbortError" ? "interrupted" : "failed";
        this.activeTurn.error = err.message;
        this.activeTurn.completedAt = Date.now();

        if (this.activeTurn.status === "interrupted") {
          turnCallbacks.onEvent({ type: "turn_interrupted", turnId: this.activeTurn.turnId });
        } else {
          turnCallbacks.onEvent({
            type: "turn_failed",
            turnId: this.activeTurn.turnId,
            error: err.message,
          });
        }

        this.activeThread?.turns.push(this.activeTurn);
        if (this.activeThread) {
          this.activeThread.metadata.updatedAt = this.activeTurn.completedAt ?? Date.now();
          this.activeThread.metadata.lastTurnStatus = this.activeTurn.status;
          this.activeThread.metadata.activeTurnId = undefined;
          await this.persistThreadSnapshot(this.activeThread);
        }
      }
      throw err;
    } finally {
      this.isRunning = false;
      this.abortController = null;
      if (this.activeThread) {
        this.threadStateManager.markIdle(this.activeThread.metadata.threadId);
        this.publishThreadStatus();
        this.scheduleIdleThreadUnload();
        await this.persistThreadRuntime(this.activeThread.metadata.threadId);
      }
      if (this.autoDrainInputQueue && this.inputQueue.size() > 0) {
        this.scheduleInputQueueDrain();
      }
      // 通知 interrupt() 等待者：Turn 清理已完成
      this._resolveTurnCompletion?.();
      this._turnCompletionPromise = null;
      this._resolveTurnCompletion = null;
    }
  }

  /** 中断当前 Turn，等待清理完成后返回 */
  async interrupt(requestId: ConnectionRequestId = `interrupt-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`): Promise<void> {
    this.pendingInterruptQueue.push(requestId);
    this.autoDrainInputQueue = false;
    this.inputQueue.clear();
    this.abortController?.abort();
    // 等待 runTurn() 的 catch/finally 完全执行完毕，
    // 确保 isRunning=false、activeThread/activeTurn 稳定后才返回。
    try {
      if (this._turnCompletionPromise) {
        await this._turnCompletionPromise;
      }
    } finally {
      this.inputQueue.clear();
      this.pendingInterruptQueue.drain();
    }
  }

  private scheduleInputQueueDrain(): void {
    if (!this.autoDrainInputQueue || this.isDrainingInputQueue || this.isRunning) return;
    this.isDrainingInputQueue = true;
    queueMicrotask(() => {
      void this.drainInputQueue();
    });
  }

  private async drainInputQueue(): Promise<void> {
    try {
      while (!this.isRunning) {
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
      if (this.autoDrainInputQueue && !this.isRunning && this.inputQueue.size() > 0) {
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
    if (this.isRunning || !this.activeThread) return;

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

      // 1. 构建消息历史
      const messages = turnItemGroupsToChatMessages(this.getTurnItemGroups());

      // 解析图片附件：将本地文件路径转为 base64 data URI
      await resolveImageAttachments(messages);

      // 注入中断恢复上下文（作为系统消息，不显示给用户）
      if (resumeContext) {
        messages.push({
          role: "system",
          content: resumeContext,
        });
      }

      // 2. 调用 AI（流式）
      const configuredMode: ReasoningMode = this.config.aiConfig.reasoningMode || this.config.reasoningMode || "high";

      // 构建系统提示词：静态基础 + 动态文件夹上下文
      let effectiveSystemPrompt = await buildEffectiveSystemPrompt(
        this.config.systemPrompt,
        this.activeThread?.metadata.folderId,
        {
          content: input.content,
          attachments: input.attachments,
        }
      );
      effectiveSystemPrompt = await appendRuntimeLongTermMemoryContext(
        effectiveSystemPrompt,
        this.stateRuntimeStore
      );

      const toolDefs = getToolDefs(this.config.toolExecutors);
      const streamParams = {
        messages,
        tools: toolDefs,
        systemPrompt: effectiveSystemPrompt,
        maxTokens: resolveMaxTokens(this.config.aiConfig),
        reasoningMode: configuredMode,
        signal: this.abortController?.signal,
        roundId: round,
      };

      // 3. 收集响应
      let streamResult;
      try {
        streamResult = await runAIRequestWithRetry({
          phase: "sampling",
          config: this.config.aiRequestRetryConfig?.sampling,
          signal: this.abortController?.signal,
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
        if (this.activeThread) {
          this.activeThread.metadata.totalTokenUsage = this.activeThread.metadata.totalTokenUsage
            ? mergeTokenUsage(this.activeThread.metadata.totalTokenUsage, streamResult.usage)
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
    if (!this.abortController?.signal.aborted) return;
    const error = new Error("aborted");
    error.name = "AbortError";
    throw error;
  }

  private getSessionCompactionConfig(thread = this.activeThread): CompactionConfig {
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
      signal: this.abortController?.signal,
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
        this.compactedHistory = history;
      },
      getActiveThread: () => this.activeThread,
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
      activeThread: this.activeThread,
      activeTurn: this.activeTurn,
      compactedHistory: this.compactedHistory,
    });
  }

  private getTurnItemGroups(): TurnItem[][] {
    return collectPromptTurnItemGroups({
      activeThread: this.activeThread,
      activeTurn: this.activeTurn,
      compactedHistory: this.compactedHistory,
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
      activeThread: this.activeThread,
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
    if (this.activeThread && previousWindow !== config.contextWindowSize) {
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
