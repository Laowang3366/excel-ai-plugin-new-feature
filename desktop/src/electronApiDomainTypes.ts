// Update / Office / AI provider / 文件夹等领域数据类型（渲染进程视角）

export type ExcelRangeExpandMode = "none" | "spill" | "currentArray" | "currentRegion";
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
export interface OfficeAutomationResult<T> {
  success: boolean;
  data?: T;
  error?: string;
}
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
export interface OfficeAutomationChange {
  kind: string;
  target?: string;
  detail: string;
}
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

export interface FolderFileInfo {
  fileName: string;
  filePath: string;
  size: number;
  lastModified: number;
}
