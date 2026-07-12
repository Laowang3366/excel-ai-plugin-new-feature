/**
 * IPC API 类型定义。
 *
 * 与 electronApi.d.ts 中 ElectronAPI 一致，供运行时 wrapper 和测试 mock 共用。
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
  SandboxConfig,
  SandboxPrefixRule,
  WindowDisplayMode,
  OfficeApplication,
  DesktopUpdateState,
  UpdateKind,
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
    launchOffice: (application: OfficeApplication) => Promise<{ success: boolean; error?: string }>;
    log: (level: string, tag: string, message: string) => Promise<void>;
  };
  update: {
    getState: () => Promise<DesktopUpdateState>;
    check: (manual?: boolean) => Promise<DesktopUpdateState>;
    download: (kind: UpdateKind) => Promise<DesktopUpdateState>;
    apply: () => Promise<DesktopUpdateState>;
    onStateChanged: (callback: (state: DesktopUpdateState) => void) => () => void;
  };
  window: {
    getAlwaysOnTop: () => Promise<boolean>;
    setAlwaysOnTop: (enabled: boolean) => Promise<boolean>;
    getDisplayMode: () => Promise<WindowDisplayMode>;
    setDisplayMode: (mode: WindowDisplayMode) => Promise<WindowDisplayMode>;
    onDisplayModeChanged: (callback: (mode: WindowDisplayMode) => void) => () => void;
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
    interrupt: (threadId?: string | null) => Promise<{ success: boolean; error?: string }>;
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
    listFilesBatch: (folderPaths: string[]) => Promise<Record<string, FolderFileInfo[]>>;
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
