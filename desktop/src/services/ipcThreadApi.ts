import type { IIpcApi } from "./ipcApiTypes";

type RawIpcApiGetter = () => IIpcApi | null;

export function createThreadIpcApi(
  getRaw: RawIpcApiGetter,
): Pick<IIpcApi, "thread" | "threadGraph"> {
  return {
    thread: {
      list: async () => {
        const raw = getRaw();
        if (!raw) return [];
        return raw.thread.list();
      },
      load: async (threadId) => {
        const raw = getRaw();
        if (!raw) throw new Error("IPC not available");
        return raw.thread.load(threadId);
      },
      delete: async (threadId) => {
        const raw = getRaw();
        if (!raw) return false;
        return raw.thread.delete(threadId);
      },
      resume: async (threadId) => {
        const raw = getRaw();
        if (!raw) return { success: false };
        return raw.thread.resume(threadId);
      },
      newThread: async (folderId) => {
        const raw = getRaw();
        if (!raw) return { success: false };
        return raw.thread.newThread(folderId);
      },
      updateMetadata: async (threadId, patch) => {
        const raw = getRaw();
        if (!raw) return;
        return raw.thread.updateMetadata(threadId, patch);
      },
      findLatest: async () => {
        const raw = getRaw();
        if (!raw) return null;
        return raw.thread.findLatest();
      },
      runtimeStatus: async () => {
        const raw = getRaw();
        if (!raw) {
          return { status: "not_loaded", idleUnloadMs: 0 };
        }
        return raw.thread.runtimeStatus();
      },
    },

    threadGraph: {
      upsertSpawnEdge: async (parentThreadId, childThreadId, label) => {
        const raw = getRaw();
        if (!raw) throw new Error("IPC not available");
        return raw.threadGraph.upsertSpawnEdge(parentThreadId, childThreadId, label);
      },
      closeSpawnEdge: async (parentThreadId, childThreadId) => {
        const raw = getRaw();
        if (!raw) return null;
        return raw.threadGraph.closeSpawnEdge(parentThreadId, childThreadId);
      },
      listDescendants: async (parentThreadId, status) => {
        const raw = getRaw();
        if (!raw) return [];
        return raw.threadGraph.listDescendants(parentThreadId, status);
      },
    },
  };
}
