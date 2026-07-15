import type {
  AgentEvent,
  FileAttachment,
  ThreadData,
  ThreadMetadata,
  ThreadRuntimeSnapshot,
  ThreadSpawnDescendant,
  ThreadSpawnEdge,
  ThreadSpawnStatusFilter,
} from "./electronApiAgentTypes";
import type {
  DesktopUpdateState,
  ExcelRangeExpandMode,
  FolderFileInfo,
  OfficeApplication,
  OfficeAutomationApp,
  OfficeAutomationDocument,
  OfficeAutomationObject,
  OfficeAutomationResult,
  OfficeAutomationTemplate,
  OfficeAutomationTransaction,
  OfficeAutomationWorkflow,
  UpdateKind,
  WindowDisplayMode,
} from "./electronApiDomainTypes";

/** electronAPI 接口 — 与 preload.ts 完全对齐 */
export interface ElectronAPI {
  app: {
    getDataPath: () => Promise<string>;
    selectDataPath: () => Promise<{ canceled: boolean; filePaths: string[] }>;
    selectExportPath: () => Promise<{ canceled: boolean; filePaths: string[] }>;
    migrateDataPath: (targetPath: string) => Promise<{
      success: boolean;
      dataPath?: string;
      error?: string;
      oldRootCleared?: boolean;
      oldRootError?: string;
    }>;
    exportUserData: (targetPath: string) => Promise<{
      success: boolean;
      exportPath?: string;
      exportedAt?: string;
      categories?: string[];
      error?: string;
    }>;
    eraseUserData: (input: { confirmation: string }) => Promise<{
      success: boolean;
      erasedCategories: string[];
      errors: string[];
      error?: string;
      proofSummary?: {
        createdAt: string;
        installIdDigest: string;
        proofDigest: string;
        destroyedKeyCount: number;
        keyMaterialDestroyed: boolean;
        replicaCount: number;
        erasedCount: number;
        failedCount: number;
      };
    }>;
    rotateLocalDataKey: () => Promise<{ success: boolean; keyId?: number; error?: string }>;
    openPath: (targetPath: string) => Promise<string>;
    openExternal: (targetUrl: string) => Promise<string>;
    launchOffice: (application: OfficeApplication) => Promise<{ success: boolean; error?: string }>;
    log: (level: string, tag: string, message: string) => Promise<void>;
  };
  update: {
    getState: () => Promise<DesktopUpdateState>;
    ackHotPatchHealth: () => Promise<boolean>;
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
    set: (key: string, value: unknown) => Promise<unknown>;
    getAll: () => Promise<Record<string, unknown>>;
  };
  excel: {
    detectStatus: () => Promise<{
      connected: boolean;
      host: string;
      version?: string;
      workbookName?: string;
      availableHosts?: string[];
    }>;
    connect: () => Promise<{
      connected: boolean;
      host: string;
      version?: string;
      workbookName?: string;
    }>;
    selectHost: (
      host: "excel" | "wps",
    ) => Promise<{ connected: boolean; host: string; version?: string; workbookName?: string }>;
    getSelection: () => Promise<{ address: string; values: unknown[][]; sheetName: string }>;
    getSelectionAddress: () => Promise<{ address: string; sheetName: string }>;
    readRange: (
      sheetName: string,
      range: string,
      expand?: ExcelRangeExpandMode,
    ) => Promise<{
      values: unknown[][];
      address?: string;
      expanded?: boolean;
      expandMode?: string;
    }>;
    inspectWorkbook: () => Promise<unknown>;
    writeRange: (
      sheetName: string,
      range: string,
      values: unknown[][],
    ) => Promise<{ success: boolean; error?: string }>;
  };
  office: {
    detectWordStatus: () => Promise<{
      connected: boolean;
      host: string;
      version?: string;
      documentName?: string;
    }>;
    detectPresentationStatus: () => Promise<{
      connected: boolean;
      host: string;
      version?: string;
      presentationName?: string;
    }>;
    automation: {
      documents: {
        list: (
          app?: OfficeAutomationApp,
        ) => Promise<OfficeAutomationResult<OfficeAutomationDocument[]>>;
        activate: (input: {
          app: OfficeAutomationApp;
          filePath: string;
          instanceId?: string;
        }) => Promise<OfficeAutomationResult<OfficeAutomationDocument>>;
      };
      objects: {
        list: (input: {
          app: OfficeAutomationApp;
          filePath: string;
          instanceId?: string;
          kind?: string;
        }) => Promise<OfficeAutomationResult<OfficeAutomationObject[]>>;
        activate: (input: {
          app: OfficeAutomationApp;
          filePath: string;
          instanceId?: string;
          locator: string;
        }) => Promise<OfficeAutomationResult<OfficeAutomationObject>>;
      };
      workflows: {
        list: () => Promise<OfficeAutomationResult<OfficeAutomationWorkflow[]>>;
        get: (id: string) => Promise<OfficeAutomationResult<OfficeAutomationWorkflow>>;
        cancel: (id: string) => Promise<OfficeAutomationResult<OfficeAutomationWorkflow>>;
        resume: (id: string) => Promise<OfficeAutomationResult<unknown>>;
      };
      templates: {
        list: () => Promise<OfficeAutomationResult<OfficeAutomationTemplate[]>>;
        saveFromWorkflow: (input: {
          workflowId: string;
          templateId?: string;
          name: string;
          description?: string;
        }) => Promise<OfficeAutomationResult<OfficeAutomationTemplate>>;
        delete: (id: string) => Promise<OfficeAutomationResult<boolean>>;
        run: (input: {
          templateId: string;
          variables?: Record<string, unknown>;
        }) => Promise<OfficeAutomationResult<unknown>>;
      };
      transactions: {
        list: () => Promise<OfficeAutomationResult<OfficeAutomationTransaction[]>>;
        get: (id: string) => Promise<OfficeAutomationResult<OfficeAutomationTransaction>>;
        undo: (
          id: string,
          force?: boolean,
        ) => Promise<OfficeAutomationResult<OfficeAutomationTransaction>>;
        redo: (
          id: string,
          force?: boolean,
        ) => Promise<OfficeAutomationResult<OfficeAutomationTransaction>>;
      };
    };
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
    }) => Promise<{
      success: boolean;
      queued?: boolean;
      queueSize?: number;
      turnId?: string;
      threadId?: string;
      error?: string;
    }>;
    interrupt: (threadId?: string | null) => Promise<{ success: boolean; error?: string }>;
    onEvent: (callback: (event: AgentEvent) => void) => () => void;
    onStreamDelta: (
      callback: (data: {
        delta: string;
        itemType: string;
        roundId?: number;
        threadId?: string;
        clientId?: string;
      }) => void,
    ) => () => void;
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
    upsertSpawnEdge: (
      parentThreadId: string,
      childThreadId: string,
      label?: string,
    ) => Promise<ThreadSpawnEdge>;
    closeSpawnEdge: (
      parentThreadId: string,
      childThreadId: string,
    ) => Promise<ThreadSpawnEdge | null>;
    listDescendants: (
      parentThreadId: string,
      status?: ThreadSpawnStatusFilter,
    ) => Promise<ThreadSpawnDescendant[]>;
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
    writeTempFile: (data: {
      prefix?: string;
      suffix?: string;
      data: string;
    }) => Promise<{ success: boolean; filePath?: string; error?: string }>;
    /** 获取拖拽/粘贴 File 对象对应的本地路径（Electron 新版替代 File.path） */
    getPathForFile?: (file: File) => string;
  };
  folder: {
    /** 列出文件夹内的 Office 文件（Excel/Word/PowerPoint） */
    listFiles: (folderPath: string) => Promise<FolderFileInfo[]>;
    /** 批量列出多个文件夹内的 Office 文件 */
    listFilesBatch: (folderPaths: string[]) => Promise<Record<string, FolderFileInfo[]>>;
  };
  tools: {
    list: () => Promise<unknown[]>;
  };
  tool: {
    /** 确认执行挂起的工具调用 */
    confirm: (toolCallId: string, alwaysAllow?: boolean) => Promise<void>;
    /** 取消挂起的工具调用 */
    cancel: (toolCallId: string) => Promise<void>;
  };
  ai: {
    /** 获取可用模型列表 */
    listModels: (
      baseUrl: string,
      apiKey: string,
      apiFormat: string,
      providerId?: string,
    ) => Promise<string[]>;
    /** 测试 API 连接 */
    testConnection: (
      baseUrl: string,
      apiKey: string,
      apiFormat: string,
      model: string,
      providerId?: string,
    ) => Promise<{ success: boolean; error?: string; latency?: number }>;
  };
  stats: {
    /** 获取聚合的使用统计 */
    getSummary: () => Promise<
      Array<{
        turnId: string;
        threadId: string;
        model: string;
        timestamp: number;
        messages: number;
        tokens: number;
        estimated: boolean;
      }>
    >;
  };
  ocr: {
    recognize: (mode: string, filePaths: string[]) => Promise<unknown | null>;
  };
  /** 知识库 (RAG) */
  knowledge: {
    /** 列出所有已索引的知识来源 */
    listSources: () => Promise<
      Array<{
        sourcePath: string;
        sourceName: string;
        sourceType: string;
        entryCount: number;
        firstIndexed: number;
        lastIndexed: number;
        fileHash: string;
      }>
    >;
    /** 搜索知识库 */
    search: (
      query: string,
      topK?: number,
    ) => Promise<{
      success: boolean;
      data?: any[];
      error?: string;
    }>;
    /** 索引单个文件 */
    indexFile: (filePath: string) => Promise<{
      sourcePath: string;
      success: boolean;
      error?: string;
      entryCount: number;
      durationMs: number;
    }>;
    /** 索引文件夹 */
    indexFolder: (folderPath: string) => Promise<any>;
    /** 删除文件索引 */
    deleteFile: (sourcePath: string) => Promise<{ success: boolean; error?: string }>;
    /** 重建全部索引 */
    reindexAll: () => Promise<{ success: boolean; error?: string; results?: any[] }>;
  };
}
