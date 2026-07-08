import type { IIpcApi } from "./ipcApiTypes";

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
      onDisplayModeChanged: () => () => {},
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
      listFilesBatch: async () => ({}),
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
    activation: {
      getStatus: async () => ({ activated: false }),
      activate: async () => ({ success: false, error: "mock not implemented" }),
      clear: async () => ({ success: true }),
      getServerUrl: async () => "http://localhost:3456",
      setServerUrl: async () => ({ success: true }),
      checkValid: async () => false,
      getMachineInfo: async () => ({ machineId: "mock-id", machineName: "mock-machine" }),
      listDevices: async () => ({ success: true, data: { machines: [] } }),
      unbindDevice: async () => ({ success: true, currentDeviceUnbound: false }),
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
