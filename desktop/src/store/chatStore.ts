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
import type { AgentEvent, FileAttachment } from "../electronApi";
import { ipcApi } from "../services/ipcApi";
import { handleAgentEvent } from "./agentEventHandler";
import { createClearedMessagesPatch, createInitialChatState } from "./chatInitialState";
import type { ChatActions, ChatState } from "./chatStoreTypes";
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
import { reconcileRunningThreadIds } from "./chatThreadRuntimeState";
import {
  interruptTurnAction,
  resumeFromInterruptionAction,
  sendMessageAction,
} from "./chatTurnActions";

export { mergeBufferedStreamDeltas, STREAM_DELTA_STORE_FLUSH_MS };
export type { StreamDeltaInput };
export type { ChatActions, ChatState } from "./chatStoreTypes";

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

let latestThreadNavigationRequestId = 0;
let pendingThreadSwitchTargetId: string | null = null;

function beginThreadNavigation(targetThreadId: string | null): number {
  pendingThreadSwitchTargetId = targetThreadId;
  return ++latestThreadNavigationRequestId;
}

function completeThreadNavigation(requestId: number): boolean {
  if (requestId !== latestThreadNavigationRequestId) {
    return false;
  }
  pendingThreadSwitchTargetId = null;
  return true;
}

function invalidateThreadNavigation(): void {
  latestThreadNavigationRequestId += 1;
  pendingThreadSwitchTargetId = null;
}

// ============================================================
// Zustand Store
// ============================================================

export const useChatStore = create<ChatState & ChatActions>((set, get) => ({
  ...createInitialChatState(),

  // ---- Actions ----

  sendMessage: (content: string, attachments?: FileAttachment[]) =>
    sendMessageAction({ set, get }, content, attachments),

  resumeFromInterruption: (content: string, attachments?: FileAttachment[]) =>
    resumeFromInterruptionAction({ set, get }, content, attachments),

  interruptTurn: () => interruptTurnAction({ set, get }),

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

  clearMessages: () => set(createClearedMessagesPatch()),

  // ---- 事件处理 ----

  handleAgentEvent: (event: AgentEvent) => {
    const current = get();
    const isPendingInterruptEvent = Boolean(
      event.type === "turn_interrupted" &&
      event.threadId &&
      current.pendingInterruptThreadIds[event.threadId],
    );
    const patches = handleAgentEvent(event, current, []);
    const merged = applyPatches(patches);
    if (Object.keys(merged).length > 0) {
      set(merged);
    }
    if (
      !isPendingInterruptEvent &&
      (event.type === "turn_completed" ||
        event.type === "turn_interrupted" ||
        event.type === "turn_failed")
    ) {
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
    const requestId = beginThreadNavigation(threadId);
    const current = get();
    const result = await switchThreadAction(threadId, current);
    if (!completeThreadNavigation(requestId)) {
      return;
    }
    if (result.error) {
      set({ error: result.error });
    } else if (result.patches.length > 0) {
      set(applyPatches(result.patches));
    }
  },

  createNewThread: async (folderId?: string) => {
    const requestId = beginThreadNavigation(null);
    const { isStreaming, runningThreadIds } = get();
    const hasRunningThread = isStreaming || Object.values(runningThreadIds).some(Boolean);
    const result = await createNewThreadAction(folderId, hasRunningThread);
    if (!completeThreadNavigation(requestId)) {
      return;
    }
    if (result.error) {
      set({ error: result.error });
    } else if (result.patches.length > 0) {
      set(applyPatches(result.patches));
    }
  },

  deleteThread: async (threadId: string) => {
    const current = get();
    const changesActiveThread =
      current.activeThreadId === threadId || pendingThreadSwitchTargetId === threadId;
    const requestId = changesActiveThread ? beginThreadNavigation(null) : null;
    const result = await deleteThreadAction(threadId, current);
    if (result.error) {
      if (requestId === null || completeThreadNavigation(requestId)) {
        set({ error: result.error });
      }
    } else {
      if (requestId !== null) {
        completeThreadNavigation(requestId);
      }

      const latest = get();
      if (pendingThreadSwitchTargetId === threadId) {
        invalidateThreadNavigation();
      }
      if (latest.activeThreadId === threadId) {
        set({
          ...createClearedMessagesPatch(),
          activeThreadId: null,
          activeClientId: null,
          isStreaming: false,
          pendingFolderId: null,
        });
      }
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
