import type { FileAttachment, ThreadMetadata } from "../electronApi";
import { ipcApi } from "../services/ipcApi";
import type { ChatActions, ChatState } from "./chatStore";
import { buildTurnStartPatch } from "./chatTurnState";

type ChatSet = (patch: Partial<ChatState> | ((state: ChatState) => Partial<ChatState>)) => void;
type ChatGet = () => ChatState & Pick<ChatActions, "loadThreads">;

interface TurnActionContext {
  set: ChatSet;
  get: ChatGet;
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

export async function sendMessageAction(
  { set, get }: TurnActionContext,
  content: string,
  attachments?: FileAttachment[]
) {
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
}

export async function resumeFromInterruptionAction(
  { set, get }: TurnActionContext,
  content: string,
  attachments?: FileAttachment[]
) {
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
}

export function interruptTurnAction({ set, get }: TurnActionContext) {
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
}
