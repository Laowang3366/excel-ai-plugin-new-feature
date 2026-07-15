// Agent / Thread / Turn 前端投影类型（渲染进程视角）

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

export interface TokenUsage {
  inputTokens: number;
  outputTokens: number;
  cachedInputTokens?: number;
  reasoningOutputTokens?: number;
}

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

export interface FileAttachment {
  filePath: string;
  fileName: string;
  fileType: "image" | "document";
  size?: number;
}
