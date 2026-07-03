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

import type {
  AgentEvent,
  FileAttachment,
  FolderFileInfo,
  ThreadMetadata,
  ThreadData,
  ThreadRuntimeSnapshot,
  ThreadSpawnDescendant,
  ThreadSpawnEdge,
  ThreadSpawnStatusFilter,
  ExcelRangeExpandMode,
  AiProviderConfig,
  TokenUsage,
  TurnItem,
  SandboxConfig,
  SandboxPrefixRule,
  WindowDisplayMode,
} from "../electronApi";

// ============================================================
// 类型定义
// ============================================================

/** 与 electronApi.d.ts 中 ElectronAPI 一致的接口 */
export interface IIpcApi {
  app: {
    getDataPath: () => Promise<string>;
    selectDataPath: () => Promise<{ canceled: boolean; filePaths: string[] }>;
    migrateDataPath: (targetPath: string) => Promise<{ success: boolean; dataPath?: string; error?: string }>;
    openPath: (targetPath: string) => Promise<string>;
    openExternal: (targetUrl: string) => Promise<string>;
  };
  window: {
    getAlwaysOnTop: () => Promise<boolean>;
    setAlwaysOnTop: (enabled: boolean) => Promise<boolean>;
    getDisplayMode: () => Promise<WindowDisplayMode>;
    setDisplayMode: (mode: WindowDisplayMode) => Promise<WindowDisplayMode>;
  };
  settings: {
    get: (key: string) => Promise<unknown>;
    set: (key: string, value: unknown) => Promise<void>;
    getAll: () => Promise<Record<string, unknown>>;
  };
  excel: {
    detectStatus: () => Promise<{ connected: boolean; host: string; version?: string; workbookName?: string; availableHosts?: string[] }>;
    connect: () => Promise<{ connected: boolean; host: string; version?: string; workbookName?: string; availableHosts?: string[] }>;
    /** 当 Office + WPS 同时运行时，用户选择目标宿主 */
    selectHost: (host: "excel" | "wps") => Promise<{ connected: boolean; host: string; version?: string; workbookName?: string }>;
    getSelection: () => Promise<{ address: string; values: unknown[][]; sheetName: string }>;
    getSelectionAddress: () => Promise<{ address: string; sheetName: string }>;
    readRange: (
      sheetName: string,
      range: string,
      expand?: ExcelRangeExpandMode
    ) => Promise<{ values: unknown[][]; address?: string; expanded?: boolean; expandMode?: string }>;
    inspectWorkbook: () => Promise<unknown>;
    writeRange: (sheetName: string, range: string, values: unknown[][]) => Promise<{ success: boolean; error?: string }>;
  };
  office: {
    detectWordStatus: () => Promise<{ connected: boolean; host: string; version?: string; documentName?: string }>;
    detectPresentationStatus: () => Promise<{ connected: boolean; host: string; version?: string; presentationName?: string }>;
  };
  agent: {
    startTurn: (input: {
      content: string;
      attachments?: FileAttachment[];
      clientId?: string;
      threadId?: string | null;
      isResume?: boolean;
      resumeContext?: string;
    }) => Promise<{ success: boolean; turnId?: string; threadId?: string; error?: string }>;
    continueTurn: (input: {
      content: string;
      attachments?: FileAttachment[];
      clientId?: string;
      threadId?: string | null;
    }) => Promise<{ success: boolean; turnId?: string; threadId?: string; error?: string }>;
    enqueueTurn: (input: {
      content: string;
      attachments?: FileAttachment[];
      clientId?: string;
      threadId?: string | null;
      isResume?: boolean;
    }) => Promise<{ success: boolean; queued?: boolean; queueSize?: number; turnId?: string; threadId?: string; error?: string }>;
    interrupt: (threadId?: string | null) => Promise<{ success: boolean }>;
    onEvent: (callback: (event: AgentEvent) => void) => () => void;
    onStreamDelta: (callback: (data: { delta: string; itemType: string; roundId?: number; threadId?: string; clientId?: string }) => void) => () => void;
  };
  thread: {
    list: () => Promise<ThreadMetadata[]>;
    load: (threadId: string) => Promise<ThreadData>;
    delete: (threadId: string) => Promise<boolean>;
    resume: (threadId: string) => Promise<{ success: boolean }>;
    newThread: (folderId?: string) => Promise<{ success: boolean }>;
    updateMetadata: (threadId: string, patch: Record<string, unknown>) => Promise<void>;
    findLatest: () => Promise<string | null>;
    runtimeStatus: () => Promise<ThreadRuntimeSnapshot>;
  };
  threadGraph: {
    upsertSpawnEdge: (parentThreadId: string, childThreadId: string, label?: string) => Promise<ThreadSpawnEdge>;
    closeSpawnEdge: (parentThreadId: string, childThreadId: string) => Promise<ThreadSpawnEdge | null>;
    listDescendants: (parentThreadId: string, status?: ThreadSpawnStatusFilter) => Promise<ThreadSpawnDescendant[]>;
  };
  dialog: {
    openFile: () => Promise<{ canceled: boolean; filePaths: string[] }>;
    openImage: () => Promise<{ canceled: boolean; filePaths: string[] }>;
    openFolder: () => Promise<{ canceled: boolean; filePaths: string[] }>;
  };
  file: {
    readAsBase64: (filePath: string) => Promise<{
      data?: string;
      mimeType?: string;
      fileName?: string;
      size?: number;
      error?: string;
    }>;
    /** 移动文件到系统回收站/废纸篓 */
    trashFile: (filePath: string) => Promise<{ success: boolean; error?: string }>;
    /** 用系统默认应用打开文件 */
    openFile: (filePath: string) => Promise<{ success: boolean; error?: string }>;
    /** 将文件绝对路径复制到剪贴板 */
    copyPath: (filePath: string) => Promise<{ success: boolean }>;
    /** 在系统文件管理器中显示文件 */
    revealInExplorer: (filePath: string) => Promise<{ success: boolean; error?: string }>;
    /** 将 base64 数据写入临时文件（截图粘贴等） */
    writeTempFile: (data: { prefix?: string; suffix?: string; data: string }) => Promise<{ success: boolean; filePath?: string; error?: string }>;
    /** 获取拖拽/粘贴 File 对象对应的本地路径（Electron 新版替代 File.path） */
    getPathForFile?: (file: File) => string;
  };
  folder: {
    listFiles: (folderPath: string) => Promise<FolderFileInfo[]>;
  };
  tools: {
    list: () => Promise<unknown[]>;
  };
  sandbox: {
    getConfig: () => Promise<SandboxConfig>;
    setUserRules: (rules: SandboxPrefixRule[]) => Promise<{ success: boolean; error?: string }>;
    setWritableRoots: (roots: string[]) => Promise<{ success: boolean; error?: string }>;
  };
  tool: {
    confirm: (toolCallId: string, alwaysAllow?: boolean) => Promise<void>;
    cancel: (toolCallId: string) => Promise<void>;
  };
  ai: {
    listModels: (baseUrl: string, apiKey: string, apiFormat: string) => Promise<string[]>;
    testConnection: (baseUrl: string, apiKey: string, apiFormat: string, model: string) => Promise<{ success: boolean; error?: string; latency?: number }>;
  };
  stats: {
    getSummary: () => Promise<Array<{
      turnId: string;
      threadId: string;
      model: string;
      timestamp: number;
      messages: number;
      tokens: number;
      estimated: boolean;
    }>>;
  };
  /** OCR 视觉识别：通过当前模型 API 的多模态能力识别图片 */
  ocr: {
    recognize: (mode: string, filePaths: string[]) => Promise<unknown | null>;
  };
  /** 知识库 (RAG) */
  knowledge: {
    listSources: () => Promise<Array<{
      sourcePath: string;
      sourceName: string;
      sourceType: string;
      entryCount: number;
      firstIndexed: number;
      lastIndexed: number;
      fileHash: string;
    }>>;
    search: (query: string, topK?: number) => Promise<{ success: boolean; data?: any[]; error?: string }>;
    indexFile: (filePath: string) => Promise<{ sourcePath: string; success: boolean; error?: string; entryCount: number; durationMs: number }>;
    indexFolder: (folderPath: string) => Promise<any>;
    deleteFile: (sourcePath: string) => Promise<{ success: boolean; error?: string }>;
    reindexAll: () => Promise<{ success: boolean; error?: string; results?: any[] }>;
  };
}

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

  knowledge: {
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
      if (!raw) return { sourcePath: filePath, success: false, error: "IPC not available", entryCount: 0, durationMs: 0 };
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
  },
};

// ============================================================
// 测试辅助
// ============================================================

/**
 * 创建部分 mock 的 IPC API 实例。
 *
 * 用法：
 *   vi.mock("../services/ipcApi", () => ({
 *     ipcApi: createMockIpcApi({
 *       settings: { getAll: vi.fn().mockResolvedValue({...}) },
 *     }),
 *   }));
 */
export function createMockIpcApi(overrides: Partial<IIpcApi> = {}): IIpcApi {
  // 深合并：只覆盖提供的部分，其余使用默认空值
  const defaults: IIpcApi = {
    app: {
      getDataPath: async () => "/mock/data",
      selectDataPath: async () => ({ canceled: true, filePaths: [] }),
      migrateDataPath: async () => ({ success: true }),
      openPath: async () => "",
      openExternal: async () => "",
    },
    window: {
      getAlwaysOnTop: async () => false,
      setAlwaysOnTop: async () => false,
      getDisplayMode: async () => "normal",
      setDisplayMode: async (mode) => mode,
    },
    settings: {
      get: async () => undefined,
      set: async () => {},
      getAll: async () => ({}),
    },
    excel: {
      detectStatus: async () => ({ connected: false, host: "" }),
      connect: async () => ({ connected: false, host: "" }),
      selectHost: async () => ({ connected: false, host: "" }),
      getSelection: async () => ({ address: "", values: [], sheetName: "" }),
      getSelectionAddress: async () => ({ address: "", sheetName: "" }),
      readRange: async () => ({ values: [] }),
      inspectWorkbook: async () => null,
      writeRange: async () => ({ success: false, error: "not implemented" }),
    },
    office: {
      detectWordStatus: async () => ({ connected: false, host: "unknown" }),
      detectPresentationStatus: async () => ({ connected: false, host: "unknown" }),
    },
    agent: {
      startTurn: async () => ({ success: false, error: "not implemented" }),
      continueTurn: async () => ({ success: false, error: "not implemented" }),
      enqueueTurn: async () => ({ success: false, error: "not implemented" }),
      interrupt: async (_threadId?: string | null) => ({ success: false }),
      onEvent: () => () => {},
      onStreamDelta: () => () => {},
    },
    thread: {
      list: async () => [],
      load: async () => { throw new Error("not implemented"); },
      delete: async () => false,
      resume: async () => ({ success: false }),
      newThread: async () => ({ success: false }),
      updateMetadata: async () => {},
      findLatest: async () => null,
      runtimeStatus: async () => ({ status: "not_loaded", idleUnloadMs: 0 }),
    },
    threadGraph: {
      upsertSpawnEdge: async () => ({
        parentThreadId: "",
        childThreadId: "",
        status: "open",
        createdAt: 0,
      }),
      closeSpawnEdge: async () => null,
      listDescendants: async () => [],
    },
    dialog: {
      openFile: async () => ({ canceled: true, filePaths: [] }),
      openImage: async () => ({ canceled: true, filePaths: [] }),
      openFolder: async () => ({ canceled: true, filePaths: [] }),
    },
    file: {
      readAsBase64: async () => ({ error: "not implemented" }),
      trashFile: async () => ({ success: false, error: "not implemented" }),
      openFile: async () => ({ success: false, error: "not implemented" }),
      copyPath: async () => ({ success: false }),
      revealInExplorer: async () => ({ success: false, error: "not implemented" }),
      writeTempFile: async () => ({ success: false, error: "not implemented" }),
      getPathForFile: () => "",
    },
    folder: {
      listFiles: async () => [],
    },
    tools: {
      list: async () => [],
    },
    sandbox: {
      getConfig: async () => ({ defaultRules: [], userRules: [], extraWritableRoots: [] }),
      setUserRules: async () => ({ success: false, error: "not implemented" }),
      setWritableRoots: async () => ({ success: false, error: "not implemented" }),
    },
    tool: {
      confirm: async () => {},
      cancel: async () => {},
    },
    ai: {
      listModels: async () => [],
      testConnection: async () => ({ success: false, error: "not implemented" }),
    },
    stats: {
      getSummary: async () => [],
    },
    ocr: {
      recognize: async () => null,
    },
    knowledge: {
      listSources: async () => [],
      search: async () => ({ success: false, error: "not implemented" }),
      indexFile: async () => ({ sourcePath: "", success: false, error: "not implemented", entryCount: 0, durationMs: 0 }),
      indexFolder: async () => ({ success: false, error: "not implemented" }),
      deleteFile: async () => ({ success: false, error: "not implemented" }),
      reindexAll: async () => ({ success: false, error: "not implemented" }),
    },
  };

  // 浅合并每个命名空间
  const result = { ...defaults };
  for (const [namespace, impl] of Object.entries(overrides)) {
    if (impl && typeof impl === "object") {
      (result as any)[namespace] = { ...(result as any)[namespace], ...impl };
    }
  }
  return result;
}
