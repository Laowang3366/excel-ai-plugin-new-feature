/**
 * threadActions - 会话管理操作
 */

import { ipcApi } from "../services/ipcApi";
import type { ThreadMetadata } from "../electronApi";
import type { ChatState } from "./chatStore";

export async function loadThreads(): Promise<{ threads: ThreadMetadata[] }> {
  try {
    const threadList = await ipcApi.thread.list();
    return { threads: (threadList as ThreadMetadata[]) || [] };
  } catch {
    return { threads: [] };
  }
}

function isThreadRunning(threadId: string, current: ChatState): boolean {
  if (current.stoppedThreadIds?.[threadId]) return false;
  const metadata = current.threads?.find((thread) => thread.threadId === threadId);
  return Boolean(
    current.runningThreadIds?.[threadId] ||
    metadata?.activeTurnId ||
    metadata?.lastTurnStatus === "in_progress"
  );
}

function isMetadataRunning(metadata: ThreadMetadata | undefined): boolean {
  return Boolean(metadata?.activeTurnId || metadata?.lastTurnStatus === "in_progress");
}

export async function switchThread(
  threadId: string,
  current: ChatState
): Promise<{
  patches: Array<Partial<ChatState>>;
  error?: string;
}> {
  try {
    if (!ipcApi?.thread) return { patches: [], error: "Thread API not available" };

    const threadData = await ipcApi.thread.load(threadId) as any;
    let allItems: any[] = [];
    if (threadData?.turns && Array.isArray(threadData.turns)) {
      allItems = threadData.turns.flatMap((t: any) => t.items || []);
    } else if (threadData?.items && Array.isArray(threadData.items)) {
      allItems = threadData.items;
    }

    const loadedMetadata = threadData?.metadata as ThreadMetadata | undefined;
    const cachedMetadata = current.threads?.find((thread) => thread.threadId === threadId);
    const targetMetadata = loadedMetadata ?? cachedMetadata;
    const targetRunning = !current.stoppedThreadIds?.[threadId] && (
      Boolean(current.runningThreadIds?.[threadId]) ||
      isMetadataRunning(loadedMetadata) ||
      (!loadedMetadata && isThreadRunning(threadId, current))
    );

    return {
      patches: [{
        activeThreadId: threadId,
        activeClientId: null,
        messages: allItems,
        isStreaming: targetRunning,
        streamingContent: "",
        streamingReasoning: "",
        activeStreamingRound: null,
        activeTurnId: targetRunning ? targetMetadata?.activeTurnId ?? null : null,
        turnStatus: targetRunning ? "in_progress" : "idle",
        lastInterruptContext: null,
        compactionNotice: null,
        error: null,
        pendingFolderId: null,
      }],
    };
  } catch (err: any) {
    return { patches: [], error: `切换会话失败：${err.message}` };
  }
}

export async function createNewThread(
  folderId?: string,
  isCurrentlyStreaming?: boolean
): Promise<{
  patches: Array<Partial<ChatState>>;
  error?: string;
}> {
  try {
    if (isCurrentlyStreaming) {
      return {
        patches: [],
        error: "当前会话正在执行，请等待完成或停止后再新建会话",
      };
    }

    await ipcApi.thread.newThread(folderId);

    return {
      patches: [{
        messages: [],
        isStreaming: false,
        streamingContent: "",
        streamingReasoning: "",
        activeStreamingRound: null,
        activeThreadId: null,
        activeTurnId: null,
        activeClientId: null,
        turnStatus: "idle",
        lastInterruptContext: null,
        tokenUsage: null,
        contextUsage: null,
        compactionNotice: null,
        error: null,
        pendingFolderId: folderId || null,
      }],
    };
  } catch (err: any) {
    return { patches: [], error: `新建会话失败：${err.message}` };
  }
}

export async function deleteThread(threadId: string, current: ChatState): Promise<{
  patches: Array<Partial<ChatState>>;
  error?: string;
}> {
  try {
    await ipcApi.thread.delete(threadId);

    const patches: Array<Partial<ChatState>> = [];
    if (current.activeThreadId === threadId) {
      patches.push({
        messages: [],
        activeThreadId: null,
        turnStatus: "idle",
        pendingFolderId: null,
      });
    }

    return { patches };
  } catch (err: any) {
    return { patches: [], error: `删除会话失败：${err.message}` };
  }
}

export async function moveThreadToFolder(threadId: string, folderId?: string): Promise<{
  error?: string;
}> {
  try {
    await ipcApi.thread.updateMetadata(threadId, { folderId });
    return {};
  } catch (err: any) {
    return { error: `移动会话失败：${err.message}` };
  }
}
