import type { IIpcApi } from "./ipcApiTypes";

export function createKnowledgeIpcApi(getRaw: () => IIpcApi | null): IIpcApi["knowledge"] {
  return {
    listSources: async () => {
      const raw = getRaw();
      if (!raw) return [];
      return raw.knowledge.listSources();
    },
    search: async (query, topK) => {
      const raw = getRaw();
      if (!raw) return { success: false, error: "IPC not available" };
      return raw.knowledge.search(query, topK);
    },
    indexFile: async (filePath) => {
      const raw = getRaw();
      if (!raw) {
        return {
          sourcePath: filePath,
          success: false,
          error: "IPC not available",
          entryCount: 0,
          durationMs: 0,
        };
      }
      return raw.knowledge.indexFile(filePath);
    },
    indexFolder: async (folderPath) => {
      const raw = getRaw();
      if (!raw) return { success: false, error: "IPC not available" };
      return raw.knowledge.indexFolder(folderPath);
    },
    deleteFile: async (sourcePath) => {
      const raw = getRaw();
      if (!raw) return { success: false, error: "IPC not available" };
      return raw.knowledge.deleteFile(sourcePath);
    },
    reindexAll: async () => {
      const raw = getRaw();
      if (!raw) return { success: false, error: "IPC not available" };
      return raw.knowledge.reindexAll();
    },
  };
}
