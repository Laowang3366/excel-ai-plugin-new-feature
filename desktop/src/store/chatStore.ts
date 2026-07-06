/**
 * 聊天状态管理 — 参考 Codex 事件驱动模型重写
 *
 * 核心原则（来自 Codex）：
 * - 消息列表是 Agent 事件的 **唯一投影**，前端不自行创建消息
 * - 所有 TurnItem 通过 item_started/item_completed 事件从 Electron 主进程到达
 * - streamingContent/streamingReasoning 仅用于流式显示，不固化为消息
 * - turn_completed 只清理流式状态，不再从 streamingContent 创建新消息
 *   （因为 item_completed 已经添加了完整的 assistant_message）
 *
 * 已拆分模块：
 * - store/agentEventHandler.ts: Agent 事件 → 状态 patches
 * - store/threadActions.ts: 会话管理（加载/切换/新建/删除/移动）
 */

import { create } from "zustand";
import type { TurnItem, AgentEvent, TokenUsage, ThreadMetadata, FileAttachment } from "../electronApi";
import { ipcApi } from "../services/ipcApi";
import { handleAgentEvent } from "./agentEventHandler";
import {
  mergeBufferedStreamDeltas,
  setupChatStreamListeners,
  STREAM_DELTA_STORE_FLUSH_MS,
  type StreamDeltaInput,
} from "./chatStreamBuffer";
import {
  loadThreads as loadThreadsAction,
  switchThread as switchThreadAction,
  createNewThread as createNewThreadAction,
  deleteThread as deleteThreadAction,
  moveThreadToFolder as moveThreadToFolderAction,
} from "./threadActions";
import { buildTurnStartPatch } from "./chatTurnState";
import { reconcileRunningThreadIds } from "./chatThreadRuntimeState";

export { mergeBufferedStreamDeltas, STREAM_DELTA_STORE_FLUSH_MS };
export type { StreamDeltaInput };

// ============================================================
// 状态类型
// ============================================================

export interface ChatState {
  /** 当前 Turn 的消息列表（只从 Agent 事件产出） */
  messages: TurnItem[];
  /** 是否正在流式输出 */
  isStreaming: boolean;
  /** 流式输出的增量内容（助手消息，仅用于实时展示） */
  streamingContent: string;
  /** 流式输出的推理增量（仅用于实时展示） */
  streamingReasoning: string;
  /**
   * 当前流式轮次 ID。
   * 当新 round 的 streamDelta 到达时，与上一轮不同就会清空
   * streamingContent/streamingReasoning，防止跨轮内容泄漏。
   */
  activeStreamingRound: number | null;
  /** 是否显示推理过程 */
  showReasoning: boolean;
  /** 推理面板的展开状态（按消息 ID） */
  reasoningExpanded: Record<string, boolean>;
  /** 当前 Turn ID */
  activeTurnId: string | null;
  /** 当前活跃的会话 ID */
  activeThreadId: string | null;
  /** 当前前端发起但尚未绑定 threadId 的请求 ID */
  activeClientId: string | null;
  /** 正在执行中的会话 ID 映射，用于切换会话后仍展示运行态 */
  runningThreadIds: Record<string, boolean>;
  /** 用户明确点击停止的会话 ID，用于屏蔽旧快照里的 in_progress 状态 */
  stoppedThreadIds: Record<string, boolean>;
  /** Turn 状态 */
  turnStatus: "idle" | "in_progress" | "completed" | "interrupted" | "failed";
  /** 上次中断的上下文提示 */
  lastInterruptContext: string | null;
  /** Token 使用量 */
  tokenUsage: TokenUsage | null;
  /** 上下文使用情况 */
  contextUsage: {
    estimatedTokens: number;
    threshold: number;
    percentage: number;
    contextWindowSize: number;
  } | null;
  /** 压缩提示 */
  compactionNotice: string | null;
  /** 错误信息 */
  error: string | null;

  // ---- 会话管理 ----
  /** 所有会话列表 */
  threads: ThreadMetadata[];

  // ---- 工具审批 ----
  /** 挂起等待审批的工具调用 */
  pendingToolCall: {
    id: string;
    toolName: string;
    arguments: Record<string, unknown>;
    riskLevel: "safe" | "moderate" | "dangerous";
    description?: string;
    /** 沙箱策略给出的理由（命中 prompt 规则时） */
    sandboxJustification?: string;
  } | null;

  // ---- 输入框文件桥接 ----
  /** 待添加到输入框的文件列表 */
  pendingComposerFiles: FileAttachment[];

  /** 新建会话时暂存的文件夹 ID */
  pendingFolderId: string | null;
}

export interface ChatActions {
  /** 发送消息 */
  sendMessage: (content: string, attachments?: FileAttachment[]) => Promise<void>;
  /** 从中断处继续 */
  resumeFromInterruption: (content: string, attachments?: FileAttachment[]) => Promise<void>;
  /** 中断当前 Turn */
  interruptTurn: () => void;
  /** 切换推理显示 */
  toggleReasoning: (itemId?: string) => void;
  /** 清除错误 */
  clearError: () => void;
  /** 清空消息 */
  clearMessages: () => void;
  /** 处理 Agent 事件 */
  handleAgentEvent: (event: AgentEvent) => void;
  /** 处理流式增量 */
  handleStreamDelta: (data: StreamDeltaInput) => void;

  // ---- 会话管理 ----
  /** 加载会话列表 */
  loadThreads: () => Promise<void>;
  /** 切换到某个会话 */
  switchThread: (threadId: string) => Promise<void>;
  /** 新建会话 */
  createNewThread: (folderId?: string) => Promise<void>;
  /** 删除会话 */
  deleteThread: (threadId: string) => Promise<void>;
  /** 移动会话到文件夹 */
  moveThreadToFolder: (threadId: string, folderId?: string) => Promise<void>;

  // ---- 工具审批 ----
  /** 确认执行挂起的工具调用 */
  confirmToolCall: (toolCallId: string, alwaysAllow?: boolean) => void;
  /** 取消挂起的工具调用 */
  cancelToolCall: (toolCallId: string) => void;

  // ---- 输入框文件桥接 ----
  /** 将文件推入输入框 */
  addFilesToComposer: (files: FileAttachment[]) => void;
  /** 消费并清空待添加文件 */
  consumePendingFiles: () => FileAttachment[];
}

// ============================================================
// 事件监听器清理
// ============================================================

function setupListeners() {
  setupChatStreamListeners({
    handleAgentEvent: (event) => useChatStore.getState().handleAgentEvent(event),
    handleStreamDelta: (data) => useChatStore.getState().handleStreamDelta(data),
  });
}

// ============================================================
// 辅助：应用 patches
// ============================================================

function applyPatches(patches: Array<Partial<ChatState>>): Partial<ChatState> {
  return Object.assign({}, ...patches.filter((p) => Object.keys(p).length > 0));
}

function markThreadStopped(
  threads: ThreadMetadata[],
  threadId: string | null,
  status: "interrupted" | "failed" | "completed" = "interrupted"
): ThreadMetadata[] {
  if (!threadId) return threads;
  return threads.map((thread) =>
    thread.threadId === threadId
      ? {
          ...thread,
          activeTurnId: undefined,
          lastTurnStatus: status,
          updatedAt: Date.now(),
        }
      : thread
  );
}

async function ensureAgentThread(threadId: string | null): Promise<boolean> {
  if (!threadId) return true;
  const result = await ipcApi.thread.resume(threadId);
  return Boolean(result?.success);
}

function createClientId(): string {
  return `client-${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

// ============================================================
// Zustand Store
// ============================================================

export const useChatStore = create<ChatState & ChatActions>((set, get) => ({
  // ---- 初始状态 ----
  messages: [],
  isStreaming: false,
  streamingContent: "",
  streamingReasoning: "",
  activeStreamingRound: null,
  showReasoning: true,
  reasoningExpanded: {},
  activeTurnId: null,
  activeThreadId: null,
  activeClientId: null,
  runningThreadIds: {},
  stoppedThreadIds: {},
  turnStatus: "idle",
  lastInterruptContext: null,
  tokenUsage: null,
  contextUsage: null,
  compactionNotice: null,
  error: null,
  threads: [],
  pendingToolCall: null,
  pendingComposerFiles: [],
  pendingFolderId: null,

  // ---- Actions ----

  /**
   * 发送消息
   *
   * 参考 Codex：前端不自行创建用户消息。
   * 用户消息通过 AgentLoop 的 item_completed 事件到达前端。
   */
  sendMessage: async (content: string, attachments?: FileAttachment[]) => {
    const state = get();
    const clientId = createClientId();

    if (state.isStreaming) {
      if (!state.activeThreadId) {
        set({ error: "会话正在创建中，请等待连接完成后再发送" });
        return;
      }
      try {
        const result = await ipcApi.agent.enqueueTurn({
          content,
          attachments,
          clientId,
          threadId: state.activeThreadId,
          isResume: false,
        });
        if (!result.success) {
          set({ error: result.error || "加入队列失败" });
        } else {
          set({ error: null });
        }
      } catch (err: any) {
        set({ error: err.message });
      }
      return;
    }

    const threadReady = await ensureAgentThread(state.activeThreadId);
    if (!threadReady) {
      set({ error: "当前仍有会话在执行中，请等待完成或中断后再发送" });
      return;
    }

    set(buildTurnStartPatch(state, clientId, { compactionNotice: null }));

    try {
      const result = await ipcApi.agent.startTurn({
        content,
        attachments,
        clientId,
        threadId: state.activeThreadId,
        isResume: false,
      });

      if (!result.success) {
        set({
          turnStatus: "failed",
          error: result.error || "发送失败",
          isStreaming: false,
          activeClientId: null,
        });
      } else {
        if (result.threadId) {
          set({ activeThreadId: result.threadId, activeClientId: null, pendingFolderId: null });
        }
        get().loadThreads();
      }
    } catch (err: any) {
      set({
        turnStatus: "failed",
        error: err.message,
        isStreaming: false,
        activeClientId: null,
      });
    }
  },

  resumeFromInterruption: async (content: string, attachments?: FileAttachment[]) => {
    const clientId = createClientId();
    const threadReady = await ensureAgentThread(get().activeThreadId);
    if (!threadReady) {
      set({ error: "当前仍有会话在执行中，请等待完成或中断后再继续" });
      return;
    }

    set(buildTurnStartPatch(get(), clientId, { lastInterruptContext: null }));

    try {
      const result = await ipcApi.agent.continueTurn({
        content,
        attachments,
        clientId,
        threadId: get().activeThreadId,
      });
      if (!result.success) {
        set({
          turnStatus: "failed",
          error: result.error || "恢复失败",
          isStreaming: false,
          activeClientId: null,
        });
      } else {
        if (result.threadId) {
          set({ activeThreadId: result.threadId, activeClientId: null });
        }
      }
    } catch (err: any) {
      set({
        turnStatus: "failed",
        error: err.message,
        isStreaming: false,
        activeClientId: null,
      });
    }
  },

  interruptTurn: () => {
    const activeThreadId = get().activeThreadId;
    ipcApi.agent.interrupt(activeThreadId);
    set({
      streamingContent: "",
      streamingReasoning: "",
      activeStreamingRound: null,
      isStreaming: false,
      activeTurnId: null,
      activeClientId: null,
      runningThreadIds: activeThreadId
        ? Object.fromEntries(Object.entries(get().runningThreadIds).filter(([id]) => id !== activeThreadId))
        : get().runningThreadIds,
      stoppedThreadIds: activeThreadId
        ? { ...get().stoppedThreadIds, [activeThreadId]: true }
        : get().stoppedThreadIds,
      threads: markThreadStopped(get().threads, activeThreadId, "interrupted"),
      turnStatus: "interrupted",
      lastInterruptContext: "对话已被中断，你可以继续提问让 AI 从断点恢复",
    });
    void get().loadThreads();
  },

  toggleReasoning: (itemId?: string) => {
    if (itemId) {
      set((s) => ({
        reasoningExpanded: {
          ...s.reasoningExpanded,
          [itemId]: !s.reasoningExpanded[itemId],
        },
      }));
    } else {
      set((s) => ({ showReasoning: !s.showReasoning }));
    }
  },

  clearError: () => set({ error: null }),

  clearMessages: () =>
    set({
      messages: [],
      streamingContent: "",
      streamingReasoning: "",
      activeStreamingRound: null,
      activeTurnId: null,
      turnStatus: "idle",
      lastInterruptContext: null,
      tokenUsage: null,
      contextUsage: null,
      compactionNotice: null,
      error: null,
    }),

  // ---- 事件处理 ----

  handleAgentEvent: (event: AgentEvent) => {
    const current = get();
    const patches = handleAgentEvent(event, current, []);
    const merged = applyPatches(patches);
    if (Object.keys(merged).length > 0) {
      set(merged);
    }
    if (event.type === "turn_completed" || event.type === "turn_interrupted" || event.type === "turn_failed") {
      get().loadThreads();
    }
  },

  handleStreamDelta: (data: StreamDeltaInput) => {
    const state = get();
    if (!state.isStreaming || state.turnStatus !== "in_progress") {
      return;
    }

    const activeThreadId = state.activeThreadId;
    const activeClientId = state.activeClientId;
    if (data.threadId && activeThreadId && data.threadId !== activeThreadId) {
      return;
    }
    if (data.threadId && !activeThreadId && data.clientId !== activeClientId) {
      return;
    }

    // 跨轮保护：round 变化时清空累积，防止上一轮残留与本轮拼接
    if (data.roundId !== undefined && data.roundId !== null) {
      const currentRound = get().activeStreamingRound;
      if (currentRound !== null && currentRound !== data.roundId) {
        // 新一轮开始，重置两路 buffer
        set({ streamingContent: "", streamingReasoning: "", activeStreamingRound: data.roundId });
      } else if (currentRound === null) {
        // 首次接收 delta
        set({ activeStreamingRound: data.roundId });
      }
    }
    if (data.itemType === "assistant_message") {
      set((s) => ({ streamingContent: s.streamingContent + data.delta }));
    } else if (data.itemType === "reasoning") {
      set((s) => ({ streamingReasoning: s.streamingReasoning + data.delta }));
    }
  },

  // ---- 会话管理 ----

  loadThreads: async () => {
    const result = await loadThreadsAction();
    set((state) => ({
      threads: result.threads,
      runningThreadIds: reconcileRunningThreadIds({
        threads: result.threads,
        runningThreadIds: state.runningThreadIds,
        stoppedThreadIds: state.stoppedThreadIds,
      }),
    }));
  },

  switchThread: async (threadId: string) => {
    const current = get();
    const result = await switchThreadAction(threadId, current);
    if (result.error) {
      set({ error: result.error });
    } else if (result.patches.length > 0) {
      set(applyPatches(result.patches));
    }
  },

  createNewThread: async (folderId?: string) => {
    const { isStreaming, runningThreadIds } = get();
    const hasRunningThread = isStreaming || Object.values(runningThreadIds).some(Boolean);
    const result = await createNewThreadAction(folderId, hasRunningThread);
    if (result.error) {
      set({ error: result.error });
    } else if (result.patches.length > 0) {
      set(applyPatches(result.patches));
    }
  },

  deleteThread: async (threadId: string) => {
    const current = get();
    const result = await deleteThreadAction(threadId, current);
    if (result.error) {
      set({ error: result.error });
    } else if (result.patches.length > 0) {
      set(applyPatches(result.patches));
    }
    get().loadThreads();
  },

  moveThreadToFolder: async (threadId: string, folderId?: string) => {
    const result = await moveThreadToFolderAction(threadId, folderId);
    if (result.error) {
      set({ error: result.error });
    }
    get().loadThreads();
  },

  // ---- 工具审批 ----

  confirmToolCall: (toolCallId: string, alwaysAllow?: boolean) => {
    set({ pendingToolCall: null });
    ipcApi.tool.confirm(toolCallId, alwaysAllow);
  },

  cancelToolCall: (toolCallId: string) => {
    set({ pendingToolCall: null });
    ipcApi.tool.cancel(toolCallId);
  },

  // ---- 输入框文件桥接 ----

  addFilesToComposer: (files: FileAttachment[]) => {
    set((s) => ({ pendingComposerFiles: [...s.pendingComposerFiles, ...files] }));
  },

  consumePendingFiles: () => {
    const { pendingComposerFiles } = get();
    if (pendingComposerFiles.length === 0) return [];
    set({ pendingComposerFiles: [] });
    return pendingComposerFiles;
  },
}));

// 初始化事件监听
setupListeners();
