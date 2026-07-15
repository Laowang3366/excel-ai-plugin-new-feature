/**
 * TypeScript 类型定义 — 前端使用的所有类型
 *
 * 与 electron/agent/shared/types.ts 保持同步，
 * 但这里是渲染进程使用的"前端视角"版本。
 * Agent/Thread 投影、领域数据类型与 ElectronAPI 接口分别在
 * electronApiAgentTypes.ts、electronApiDomainTypes.ts、electronApiInterface.ts。
 */

import type { ElectronAPI } from "./electronApiInterface";

export type {
  TurnStatus,
  ThreadCompactStartParams,
  AgentEvent,
  UserMessageItem,
  AssistantMessageItem,
  ReasoningItem,
  ToolCallItem,
  ToolResultItem,
  CompactedItem,
  CompactProgressItem,
  ErrorItem,
  TurnItem,
  TokenUsage,
  ThreadMetadata,
  ThreadData,
  ThreadRuntimeStatus,
  ThreadRuntimeSnapshot,
  ThreadSpawnEdgeStatus,
  ThreadSpawnStatusFilter,
  ThreadSpawnEdge,
  ThreadSpawnDescendant,
  FileAttachment,
} from "./electronApiAgentTypes";

export type {
  ExcelRangeExpandMode,
  WindowDisplayMode,
  OfficeApplication,
  UpdateKind,
  UpdatePhase,
  DesktopUpdateState,
  OfficeAutomationApp,
  OfficeAutomationResult,
  OfficeAutomationDocument,
  OfficeAutomationObject,
  OfficeAutomationStep,
  OfficeAutomationWorkflow,
  OfficeAutomationChange,
  OfficeAutomationTransaction,
  OfficeAutomationTemplate,
  ReasoningMode,
  ModelConfig,
  AiProviderConfig,
  FolderFileInfo,
} from "./electronApiDomainTypes";

export type { ElectronAPI } from "./electronApiInterface";

// ============================================================
// Window 接口扩展 — 让 TypeScript 识别 window.electronAPI
// ============================================================

declare global {
  interface Window {
    electronAPI: ElectronAPI;
  }
}

export {};
