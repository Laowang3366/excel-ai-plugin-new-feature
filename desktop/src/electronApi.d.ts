/**
 * TypeScript 类型定义 — 前端使用的所有类型
 *
 * 与 electron/agent/shared/types.ts 保持同步，
 * 但这里是渲染进程使用的"前端视角"版本。
 */

// ============================================================
// Window 接口扩展 — 让 TypeScript 识别 window.electronAPI
// ============================================================

export type ExcelRangeExpandMode = "none" | "spill" | "currentArray" | "currentRegion";

declare global {
  interface Window {
    electronAPI: ElectronAPI;
  }
}

export {};

// ============================================================
// Agent 事件类型（从主进程接收）
// ============================================================

export type TurnStatus = "in_progress" | "completed" | "interrupted" | "failed";

/** Agent 事件（从主进程通过 IPC 发送） */
type AgentEventThreadContext = { threadId?: string; clientId?: string };

export interface ThreadCompactStartParams {
  reason: string;
  itemCount: number;
  tokensBefore: number;
  tokenThreshold: number;
  contextWindowSize?: number;
  retryCount: number;
  timestamp: number;
}

export type AgentEvent =
  | ({ type: "turn_started"; turnId: string } & AgentEventThreadContext)
  | ({ type: "turn_completed"; turnId: string; usage?: TokenUsage } & AgentEventThreadContext)
  | ({ type: "turn_interrupted"; turnId: string } & AgentEventThreadContext)
  | ({ type: "turn_failed"; turnId: string; error: string } & AgentEventThreadContext)
  | ({ type: "item_started"; item: TurnItem } & AgentEventThreadContext)
  | ({ type: "item_completed"; item: TurnItem } & AgentEventThreadContext)
  | ({ type: "item_updated"; item: TurnItem } & AgentEventThreadContext)
  | ({ type: "thread_compact_started"; params: ThreadCompactStartParams } & AgentEventThreadContext)
  | ({
      type: "context_compacted";
      summary: string;
      tokensBefore: number;
      tokensAfter: number;
    } & AgentEventThreadContext)
  | ({
      type: "context_usage";
      estimatedTokens: number;
      threshold: number;
      percentage: number;
      contextWindowSize: number;
    } & AgentEventThreadContext)
  | ({
      type: "stream_delta";
      delta: string;
      itemType: "assistant_message" | "reasoning";
      roundId?: number;
    } & AgentEventThreadContext)
  | ({
      type: "tool_approval_required";
      toolCallId: string;
      toolName: string;
      arguments: Record<string, unknown>;
      riskLevel: "safe" | "moderate" | "dangerous";
      description?: string;
      canAlwaysAllow?: boolean;
    } & AgentEventThreadContext)
  | ({ type: "error"; message: string } & AgentEventThreadContext)
  | ({ type: "warning"; message: string } & AgentEventThreadContext);

// ============================================================
// TurnItem — 对话条目
// ============================================================

export interface UserMessageItem {
  type: "user_message";
  id: string;
  content: string;
  attachments?: FileAttachment[];
  clientId?: string;
  timestamp: number;
}

export interface AssistantMessageItem {
  type: "assistant_message";
  id: string;
  content: string;
  phase?: "commentary" | "final";
  timestamp: number;
}

export interface ReasoningItem {
  type: "reasoning";
  id: string;
  summaryText: string[];
  rawContent: string[];
  timestamp: number;
}

export interface ToolCallItem {
  type: "tool_call";
  id: string;
  toolName: string;
  arguments: Record<string, unknown>;
  status: "pending" | "running" | "completed" | "failed";
  timestamp: number;
}

export interface ToolResultItem {
  type: "tool_result";
  id: string;
  toolCallId: string;
  toolName: string;
  result: unknown;
  isError: boolean;
  timestamp: number;
}

export interface CompactedItem {
  type: "compacted";
  id: string;
  summary: string;
  tokensBefore: number;
  tokensAfter: number;
  reason: string;
  timestamp: number;
}

export interface CompactProgressItem {
  type: "compact_progress";
  id: string;
  reason: string;
  status: "running" | "completed" | "failed";
  message: string;
  tokensBefore?: number;
  tokensAfter?: number;
  summary?: string;
  timestamp: number;
}

export interface ErrorItem {
  type: "error";
  id: string;
  message: string;
  timestamp: number;
}

export type TurnItem =
  | UserMessageItem
  | AssistantMessageItem
  | ReasoningItem
  | ToolCallItem
  | ToolResultItem
  | CompactedItem
  | CompactProgressItem
  | ErrorItem;

// ============================================================
// Token 使用量
// ============================================================

export interface TokenUsage {
  inputTokens: number;
  outputTokens: number;
  cachedInputTokens?: number;
  reasoningOutputTokens?: number;
}

// ============================================================
// 会话元数据
// ============================================================

export interface ThreadMetadata {
  threadId: string;
  preview: string;
  name?: string;
  modelProvider: string;
  model?: string;
  contextWindowSize?: number;
  compHash?: string;
  createdAt: number;
  updatedAt: number;
  activeTurnId?: string;
  lastTurnStatus?: TurnStatus;
  totalTokenUsage?: TokenUsage;
  /** 所属文件夹路径（对应 pinnedFolders 的 path） */
  folderId?: string;
}

/** 会话完整数据（从 thread:load 返回） */
export interface ThreadData {
  threadId: string;
  preview: string;
  name?: string;
  modelProvider: string;
  model?: string;
  createdAt: number;
  updatedAt: number;
  totalTokenUsage?: TokenUsage;
  turns: Array<{
    turnId: string;
    status: string;
    items: TurnItem[];
    tokenUsage?: TokenUsage;
    startedAt?: number;
    completedAt?: number;
  }>;
}

/** Agent 内存中的线程运行态快照 */
export type ThreadRuntimeStatus = "not_loaded" | "active" | "running" | "unloaded";

export interface ThreadRuntimeSnapshot {
  status: ThreadRuntimeStatus;
  threadId?: string;
  lastActiveAt?: number;
  unloadedAt?: number;
  idleUnloadMs: number;
}

export type ThreadSpawnEdgeStatus = "open" | "closed";
export type ThreadSpawnStatusFilter = ThreadSpawnEdgeStatus | "all";
export type WindowDisplayMode = "normal" | "compact";
export type OfficeApplication = "wps" | "excel" | "word" | "powerpoint";
export type UpdateKind = "installer" | "hotPatch";
export type UpdatePhase =
  | "idle"
  | "checking"
  | "up-to-date"
  | "available"
  | "downloading"
  | "downloaded"
  | "applying"
  | "error";

export interface DesktopUpdateState {
  phase: UpdatePhase;
  currentVersion: string;
  availableVersion?: string;
  installerAvailable: boolean;
  hotPatchAvailable: boolean;
  activeHotPatchId?: string;
  downloadedKind?: UpdateKind;
  progress?: number;
  releaseNotes: string[];
  publishedAt?: string;
  error?: string;
}

export type OfficeAutomationApp = "excel" | "word" | "presentation";
export interface OfficeAutomationResult<T> { success: boolean; data?: T; error?: string }
export interface OfficeAutomationDocument {
  app: OfficeAutomationApp;
  name: string;
  fullName?: string;
  index: number;
  active: boolean;
  progId: string;
  host: "microsoft-office" | "wps" | "unknown";
  instanceId: string;
  processId?: number;
  hwnd?: number;
  readOnly?: boolean;
  saved?: boolean;
}
export interface OfficeAutomationObject {
  app: OfficeAutomationApp;
  documentPath?: string;
  instanceId?: string;
  kind: string;
  name: string;
  locator: string;
  parent?: string;
  index?: number;
  detail?: string;
  selected?: boolean;
}
export interface OfficeAutomationStep {
  app: OfficeAutomationApp;
  action: "inspect" | "edit" | "style" | "insert" | "snapshot" | "validate";
  operation: string;
  filePath?: string;
  outputPath?: string;
  target?: string;
  params?: Record<string, unknown>;
  id?: string;
}
export interface OfficeAutomationWorkflow {
  id: string;
  status: "running" | "paused" | "done" | "failed" | "cancelled";
  createdAt: string;
  updatedAt: string;
  steps: OfficeAutomationStep[];
  sourceSteps?: OfficeAutomationStep[];
  stepRecords: Array<{
    step: number;
    id?: string;
    status: "pending" | "running" | "done" | "failed" | "skipped";
    attempts?: number;
    artifacts: string[];
    result?: { summary?: string; error?: string; changes?: OfficeAutomationChange[] };
  }>;
  completedSteps: number;
  nextStep: number;
  transactionId?: string;
  error?: string;
}
export interface OfficeAutomationChange { kind: string; target?: string; detail: string }
export interface OfficeAutomationTransaction {
  id: string;
  workflowId?: string;
  status: "pending" | "applied" | "undone" | "failed" | "conflicted";
  createdAt: string;
  updatedAt: string;
  artifacts: string[];
  changes: OfficeAutomationChange[];
  conflicts?: Array<{ filePath: string; expected: "before" | "after"; reason: string }>;
  conflictBaseStatus?: "pending" | "applied" | "undone" | "failed";
  error?: string;
}
export interface OfficeAutomationTemplate {
  id: string;
  name: string;
  description?: string;
  createdAt: string;
  updatedAt: string;
  steps: OfficeAutomationStep[];
}

export interface ThreadSpawnEdge {
  parentThreadId: string;
  childThreadId: string;
  status: ThreadSpawnEdgeStatus;
  createdAt: number;
  closedAt?: number;
  label?: string;
}

export interface ThreadSpawnDescendant {
  threadId: string;
  parentThreadId: string;
  depth: number;
  edge: ThreadSpawnEdge;
}

// ============================================================
// AI 提供商配置
// ============================================================

/** 统一思考等级枚举 */
export type ReasoningMode = "off" | "low" | "medium" | "high" | "max";

export interface ModelConfig {
  /** 模型名称/ID */
  name: string;
  /** 该模型的上下文窗口大小（tokens），覆盖供应商级默认值 */
  contextWindowSize?: number;
  /** 压缩兼容性标识；相同值的模型切换可复用压缩上下文 */
  compHash?: string;
  /** 该模型的思考等级，覆盖供应商级默认值 */
  reasoningMode?: ReasoningMode;
}

export interface AiProviderConfig {
  id: string;
  name: string;
  provider: string;
  apiKey: string;
  baseUrl: string;
  model: string;
  models?: string[]; // 可用模型列表（从 /v1/models 拉取，直连供应商使用）
  modelConfigs?: ModelConfig[]; // 结构化模型列表（聚合平台使用，每个模型可独立配置）
  defaultBaseUrl?: string;
  defaultModel?: string;
  apiFormat?: string; // API 协议格式: openai|anthropic|xunfei
  customHeaders?: Record<string, string>;
  contextWindowSize?: number; // 上下文窗口大小（tokens），供应商级默认/回退值
  compHash?: string; // 供应商级压缩兼容性标识，模型级配置可覆盖
  reasoningMode?: ReasoningMode; // 思考等级，供应商级默认值
}

// ============================================================
// 文件附件类型
// ============================================================

export interface FileAttachment {
  filePath: string;
  fileName: string;
  fileType: "image" | "document";
  size?: number;
}

// ============================================================
// 文件夹文件信息
// ============================================================

export interface FolderFileInfo {
  fileName: string;
  filePath: string;
  size: number;
  lastModified: number;
}

// ============================================================
// electronAPI 接口 — 与 preload.ts 完全对齐
// ============================================================

export interface ElectronAPI {
  app: {
    getDataPath: () => Promise<string>;
    selectDataPath: () => Promise<{ canceled: boolean; filePaths: string[] }>;
    migrateDataPath: (
      targetPath: string,
    ) => Promise<{ success: boolean; dataPath?: string; error?: string }>;
    openPath: (targetPath: string) => Promise<string>;
    openExternal: (targetUrl: string) => Promise<string>;
    launchOffice: (
      application: OfficeApplication,
    ) => Promise<{ success: boolean; error?: string }>;
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
        list: (app?: OfficeAutomationApp) => Promise<OfficeAutomationResult<OfficeAutomationDocument[]>>;
        activate: (input: { app: OfficeAutomationApp; filePath: string; instanceId?: string }) => Promise<OfficeAutomationResult<OfficeAutomationDocument>>;
      };
      objects: {
        list: (input: { app: OfficeAutomationApp; filePath: string; instanceId?: string; kind?: string }) => Promise<OfficeAutomationResult<OfficeAutomationObject[]>>;
        activate: (input: { app: OfficeAutomationApp; filePath: string; instanceId?: string; locator: string }) => Promise<OfficeAutomationResult<OfficeAutomationObject>>;
      };
      workflows: {
        list: () => Promise<OfficeAutomationResult<OfficeAutomationWorkflow[]>>;
        get: (id: string) => Promise<OfficeAutomationResult<OfficeAutomationWorkflow>>;
        cancel: (id: string) => Promise<OfficeAutomationResult<OfficeAutomationWorkflow>>;
        resume: (id: string) => Promise<OfficeAutomationResult<unknown>>;
      };
      templates: {
        list: () => Promise<OfficeAutomationResult<OfficeAutomationTemplate[]>>;
        saveFromWorkflow: (input: { workflowId: string; templateId?: string; name: string; description?: string }) => Promise<OfficeAutomationResult<OfficeAutomationTemplate>>;
        delete: (id: string) => Promise<OfficeAutomationResult<boolean>>;
        run: (input: { templateId: string; variables?: Record<string, unknown> }) => Promise<OfficeAutomationResult<unknown>>;
      };
      transactions: {
        list: () => Promise<OfficeAutomationResult<OfficeAutomationTransaction[]>>;
        get: (id: string) => Promise<OfficeAutomationResult<OfficeAutomationTransaction>>;
        undo: (id: string, force?: boolean) => Promise<OfficeAutomationResult<OfficeAutomationTransaction>>;
        redo: (id: string, force?: boolean) => Promise<OfficeAutomationResult<OfficeAutomationTransaction>>;
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
    listModels: (baseUrl: string, apiKey: string, apiFormat: string, providerId?: string) => Promise<string[]>;
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
