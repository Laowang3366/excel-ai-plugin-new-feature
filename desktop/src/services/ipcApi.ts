/**
 * IPC API 抽象层
 *
 * 封装 window.electronAPI 的直接访问，提供：
 * 1. 集中管理所有 IPC 调用
 * 2. 测试时可注入 mock 实现（通过 createMockIpcApi）
 * 3. 运行时安全检查（electronAPI 可能不存在）
 *
 * 使用方式：
 *   import { ipcApi } from "../services/ipcApi";
 *   const result = await ipcApi.settings.getAll();
 */

import type { IIpcApi } from "./ipcApiTypes";
import { createKnowledgeIpcApi } from "./ipcKnowledgeApi";
export type { IIpcApi } from "./ipcApiTypes";
export { createMockIpcApi } from "./ipcApiMock";

// ============================================================
// 实现
// ============================================================

/** 获取底层 electronAPI，不存在时返回 null */
function getRaw(): IIpcApi | null {
  if (typeof window === "undefined") return null;
  return (window as any).electronAPI ?? null;
}

/**
 * 运行时 IPC API 实例。
 *
 * 所有方法在 electronAPI 不可用时返回安全的空值，
 * 调用方无需额外判空。
 */
export const ipcApi: IIpcApi = {
  app: {
    getDataPath: async () => {
      const raw = getRaw();
      if (!raw) throw new Error("IPC not available: app.getDataPath");
      return raw.app.getDataPath();
    },
    selectDataPath: async () => {
      const raw = getRaw();
      if (!raw) return { canceled: true, filePaths: [] };
      return raw.app.selectDataPath();
    },
    migrateDataPath: async (targetPath) => {
      const raw = getRaw();
      if (!raw) return { success: false, error: "IPC not available" };
      return raw.app.migrateDataPath(targetPath);
    },
    openPath: async (targetPath) => {
      const raw = getRaw();
      if (!raw) return "";
      return raw.app.openPath(targetPath);
    },
    openExternal: async (targetUrl) => {
      const raw = getRaw();
      if (!raw) return "";
      return raw.app.openExternal(targetUrl);
    },
  },

  window: {
    getAlwaysOnTop: async () => {
      const raw = getRaw();
      if (!raw) return false;
      return raw.window.getAlwaysOnTop();
    },
    setAlwaysOnTop: async (enabled) => {
      const raw = getRaw();
      if (!raw) return false;
      return raw.window.setAlwaysOnTop(enabled);
    },
    getDisplayMode: async () => {
      const raw = getRaw();
      if (!raw?.window.getDisplayMode) return "normal";
      return raw.window.getDisplayMode();
    },
    setDisplayMode: async (mode) => {
      const raw = getRaw();
      if (!raw?.window.setDisplayMode) return "normal";
      return raw.window.setDisplayMode(mode);
    },
    onDisplayModeChanged: (callback) => {
      const raw = getRaw();
      if (!raw?.window.onDisplayModeChanged) return () => {};
      return raw.window.onDisplayModeChanged(callback);
    },
  },

  settings: {
    get: async (key) => {
      const raw = getRaw();
      if (!raw) return undefined;
      return raw.settings.get(key);
    },
    set: async (key, value) => {
      const raw = getRaw();
      if (!raw) return;
      return raw.settings.set(key, value);
    },
    getAll: async () => {
      const raw = getRaw();
      if (!raw) return {};
      return raw.settings.getAll();
    },
  },

  excel: {
    detectStatus: async () => {
      const raw = getRaw();
      if (!raw) return { connected: false, host: "" };
      return raw.excel.detectStatus();
    },
    connect: async () => {
      const raw = getRaw();
      if (!raw) return { connected: false, host: "" };
      return raw.excel.connect();
    },
    selectHost: async (host) => {
      const raw = getRaw();
      if (!raw) return { connected: false, host: "" };
      return raw.excel.selectHost(host);
    },
    getSelection: async () => {
      const raw = getRaw();
      if (!raw) return { address: "", values: [], sheetName: "" };
      return raw.excel.getSelection();
    },
    getSelectionAddress: async () => {
      const raw = getRaw();
      if (!raw) return { address: "", sheetName: "" };
      if (raw.excel.getSelectionAddress) return raw.excel.getSelectionAddress();
      const selection = await raw.excel.getSelection();
      return { address: selection.address, sheetName: selection.sheetName };
    },
    readRange: async (sheetName, range, expand) => {
      const raw = getRaw();
      if (!raw) return { values: [] };
      return raw.excel.readRange(sheetName, range, expand);
    },
    inspectWorkbook: async () => {
      const raw = getRaw();
      if (!raw) return null;
      return raw.excel.inspectWorkbook();
    },
    writeRange: async (sheetName, range, values) => {
      const raw = getRaw();
      if (!raw) return { success: false, error: "IPC not available" };
      return raw.excel.writeRange(sheetName, range, values);
    },
  },

  office: {
    detectWordStatus: async () => {
      const raw = getRaw();
      if (!raw) return { connected: false, host: "unknown" };
      return raw.office.detectWordStatus();
    },
    detectPresentationStatus: async () => {
      const raw = getRaw();
      if (!raw) return { connected: false, host: "unknown" };
      return raw.office.detectPresentationStatus();
    },
  },

  agent: {
    startTurn: async (input) => {
      const raw = getRaw();
      if (!raw) return { success: false, error: "IPC not available" };
      return raw.agent.startTurn(input);
    },
    continueTurn: async (input) => {
      const raw = getRaw();
      if (!raw) return { success: false, error: "IPC not available" };
      return raw.agent.continueTurn(input);
    },
    enqueueTurn: async (input) => {
      const raw = getRaw();
      if (!raw) return { success: false, error: "IPC not available" };
      return raw.agent.enqueueTurn(input);
    },
    interrupt: async (threadId?: string | null) => {
      const raw = getRaw();
      if (!raw) return { success: false };
      return raw.agent.interrupt(threadId);
    },
    onEvent: (callback) => {
      const raw = getRaw();
      if (!raw) return () => {};
      return raw.agent.onEvent(callback);
    },
    onStreamDelta: (callback) => {
      const raw = getRaw();
      if (!raw) return () => {};
      return raw.agent.onStreamDelta(callback);
    },
  },

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

  dialog: {
    openFile: async () => {
      const raw = getRaw();
      if (!raw) return { canceled: true, filePaths: [] };
      return raw.dialog.openFile();
    },
    openImage: async () => {
      const raw = getRaw();
      if (!raw) return { canceled: true, filePaths: [] };
      return raw.dialog.openImage();
    },
    openFolder: async () => {
      const raw = getRaw();
      if (!raw) return { canceled: true, filePaths: [] };
      return raw.dialog.openFolder();
    },
  },

  file: {
    readAsBase64: async (filePath) => {
      const raw = getRaw();
      if (!raw) return { error: "IPC not available" };
      return raw.file.readAsBase64(filePath);
    },
    trashFile: async (filePath) => {
      const raw = getRaw();
      if (!raw) return { success: false, error: "IPC not available" };
      return raw.file.trashFile(filePath);
    },
    openFile: async (filePath) => {
      const raw = getRaw();
      if (!raw) return { success: false, error: "IPC not available" };
      return raw.file.openFile(filePath);
    },
    copyPath: async (filePath) => {
      const raw = getRaw();
      if (!raw) return { success: false };
      return raw.file.copyPath(filePath);
    },
    revealInExplorer: async (filePath) => {
      const raw = getRaw();
      if (!raw) return { success: false, error: "IPC not available" };
      return raw.file.revealInExplorer(filePath);
    },
    writeTempFile: async (data) => {
      const raw = getRaw();
      if (!raw) return { success: false, error: "IPC not available" };
      return raw.file.writeTempFile(data);
    },
    getPathForFile: (file) => {
      const raw = getRaw();
      if (!raw?.file.getPathForFile) return "";
      return raw.file.getPathForFile(file);
    },
  },

  folder: {
    listFiles: async (folderPath) => {
      const raw = getRaw();
      if (!raw) return [];
      return raw.folder.listFiles(folderPath);
    },
    listFilesBatch: async (folderPaths) => {
      const raw = getRaw();
      if (!raw) return {};
      if (raw.folder.listFilesBatch) return raw.folder.listFilesBatch(folderPaths);
      const entries = await Promise.all(
        folderPaths.map(async (folderPath) => {
          try {
            return [folderPath, await raw.folder.listFiles(folderPath)] as const;
          } catch {
            return [folderPath, []] as const;
          }
        })
      );
      return Object.fromEntries(entries);
    },
  },

  tools: {
    list: async () => {
      const raw = getRaw();
      if (!raw) return [];
      return raw.tools.list();
    },
  },

  sandbox: {
    getConfig: async () => {
      const raw = getRaw();
      if (!raw) return { defaultRules: [], userRules: [], extraWritableRoots: [] };
      return raw.sandbox.getConfig();
    },
    setUserRules: async (rules) => {
      const raw = getRaw();
      if (!raw) return { success: false, error: "IPC not available" };
      return raw.sandbox.setUserRules(rules);
    },
    setWritableRoots: async (roots) => {
      const raw = getRaw();
      if (!raw) return { success: false, error: "IPC not available" };
      return raw.sandbox.setWritableRoots(roots);
    },
  },

  tool: {
    confirm: async (toolCallId, alwaysAllow) => {
      const raw = getRaw();
      if (!raw) return;
      return raw.tool.confirm(toolCallId, alwaysAllow);
    },
    cancel: async (toolCallId) => {
      const raw = getRaw();
      if (!raw) return;
      return raw.tool.cancel(toolCallId);
    },
  },

  ai: {
    listModels: async (baseUrl, apiKey, apiFormat) => {
      const raw = getRaw();
      if (!raw) return [];
      return raw.ai.listModels(baseUrl, apiKey, apiFormat);
    },
    testConnection: async (baseUrl, apiKey, apiFormat, model) => {
      const raw = getRaw();
      if (!raw) return { success: false, error: "IPC not available" };
      return raw.ai.testConnection(baseUrl, apiKey, apiFormat, model);
    },
  },

  stats: {
    getSummary: async () => {
      const raw = getRaw();
      if (!raw) return [];
      return raw.stats.getSummary();
    },
  },

  ocr: {
    recognize: async (mode, filePaths) => {
      const raw = getRaw() as any;
      if (raw?.ocr?.recognize) {
        return raw.ocr.recognize(mode, filePaths);
      }
      return null;
    },
  },

  knowledge: createKnowledgeIpcApi(getRaw),
};
