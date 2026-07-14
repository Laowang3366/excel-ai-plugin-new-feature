/**
 * 核心类型定义
 *
 * 参考 Codex 的 protocol/items.rs 和 protocol/protocol.rs，
 * 适配为 TypeScript 版本，用于 Excel AI 桌面端。
 *
 * 核心概念：
 * - Thread（会话）：一次完整的对话，包含多个 Turn
 * - Turn（轮次）：用户发一条消息 → AI 回复（可能包含工具调用）→ 结束
 * - Item（条目）：Turn 中的具体内容（消息、工具调用、推理等）
 * - Rollout（记录）：持久化的会话历史，用于恢复和回放
 */

// ============================================================
// Thread ID
// ============================================================

/** 会话唯一标识 */
export type ThreadId = string;

/** 生成新的 ThreadId */
export function generateThreadId(): ThreadId {
  return `thread-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
}

/** Turn 唯一标识 */
export type TurnId = string;

export function generateTurnId(): TurnId {
  return `turn-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

// ============================================================
// 文件附件
// ============================================================

/** 文件附件（用户通过文件选择器附加的文件） */
export interface FileAttachment {
  /** 文件绝对路径 */
  filePath: string;
  /** 文件名 */
  fileName: string;
  /** 文件类型：image 或 document */
  fileType: "image" | "document";
  /** 文件大小（字节） */
  size?: number;
}

// ============================================================
// TurnItem — 对话中的具体条目（参考 Codex TurnItem）
// ============================================================

/** 用户消息 */
export interface UserMessageItem {
  type: "user_message";
  id: string;
  content: string;
  /** 用户附加的文件列表 */
  attachments?: FileAttachment[];
  /** 客户端提供的消息 ID，用于去重 */
  clientId?: string;
  timestamp: number;
}

/** AI 助手消息 */
export interface AssistantMessageItem {
  type: "assistant_message";
  id: string;
  content: string;
  /** 消息阶段：commentary=中间评论，final=最终回答 */
  phase?: "commentary" | "final";
  timestamp: number;
}

/** AI 推理/思考过程（完整展示，不隐藏） */
export interface ReasoningItem {
  type: "reasoning";
  id: string;
  /** 推理摘要文本 */
  summaryText: string[];
  /** 原始推理内容（完整展示给用户） */
  rawContent: string[];
  timestamp: number;
}

/** 工具调用请求 */
export interface ToolCallItem {
  type: "tool_call";
  id: string;
  /** 工具名称，如 "range.read", "vba.execute" */
  toolName: string;
  /** 工具调用参数 */
  arguments: Record<string, unknown>;
  /** 工具调用状态 */
  status: "pending" | "running" | "completed" | "failed";
  timestamp: number;
}

/** 工具调用结果 */
export interface ToolResultItem {
  type: "tool_result";
  id: string;
  /** 对应的 tool_call ID */
  toolCallId: string;
  /** 工具名称 */
  toolName: string;
  /** 工具返回结果 */
  result: unknown;
  /** 是否出错 */
  isError: boolean;
  timestamp: number;
}

/** 上下文压缩事件（参考 Codex CompactedItem） */
export interface CompactedItem {
  type: "compacted";
  id: string;
  /** 压缩摘要文本 */
  summary: string;
  /** 压缩前的 token 数 */
  tokensBefore: number;
  /** 压缩后的 token 数 */
  tokensAfter: number;
  /** 压缩原因 */
  reason: CompactionReason;
  timestamp: number;
}

/** 上下文压缩进度条目，用于把压缩开始/完成/失败发送给前端 */
export interface CompactProgressItem {
  type: "compact_progress";
  id: string;
  reason: CompactionReason;
  status: "running" | "completed" | "failed";
  message: string;
  tokensBefore?: number;
  tokensAfter?: number;
  summary?: string;
  timestamp: number;
}

/** 错误条目 */
export interface ErrorItem {
  type: "error";
  id: string;
  message: string;
  timestamp: number;
}

/** Turn 中所有可能的条目类型 */
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
// Turn — 一次完整的对话轮次
// ============================================================

/** Turn 状态 */
export type TurnStatus = "in_progress" | "completed" | "interrupted" | "failed";

/** 一次对话轮次（参考 Codex StoredTurn） */
export interface Turn {
  turnId: TurnId;
  threadId: ThreadId;
  status: TurnStatus;
  items: TurnItem[];
  /** Token 使用量 */
  tokenUsage?: TokenUsage;
  /** 错误信息（当 status 为 failed 时） */
  error?: string;
  /** 开始时间 */
  startedAt: number;
  /** 结束时间 */
  completedAt?: number;
}

// ============================================================
// Thread — 一次完整的对话会话
// ============================================================

/** 会话元数据（参考 Codex StoredThread） */
export interface ThreadMetadata {
  threadId: ThreadId;
  /** 用户可见的预览文本（通常是第一条用户消息） */
  preview: string;
  /** 用户自定义的会话名称 */
  name?: string;
  /** 使用的 AI 模型提供商 */
  modelProvider: string;
  /** 使用的具体模型 */
  model?: string;
  /** 会话创建时的上下文窗口大小（tokens），用于压缩阈值计算隔离 */
  contextWindowSize?: number;
  /** 会话创建时的模型压缩兼容性标识，用于模型切换压缩审计 */
  compHash?: string;
  /** 创建时间 */
  createdAt: number;
  /** 最后更新时间 */
  updatedAt: number;
  /** 当前活跃的 Turn ID */
  activeTurnId?: TurnId;
  /** 最近一轮对话状态，用于会话列表状态显示 */
  lastTurnStatus?: TurnStatus;
  /** 累计 Token 使用量 */
  totalTokenUsage?: TokenUsage;
  /** 是否已归档 */
  archivedAt?: number;
  /** 所属文件夹路径（对应 pinnedFolders 的 path） */
  folderId?: string;
  /** 最近一次压缩后的替代历史，用于会话重启时恢复上下文 */
  compactedHistory?: TurnItem[];
}

/** 完整的会话（元数据 + 历史记录） */
export interface Thread {
  metadata: ThreadMetadata;
  turns: Turn[];
}

// ============================================================
// Thread runtime — 可变运行态快照
// ============================================================

export type ThreadRuntimeStatus = "not_loaded" | "active" | "running" | "unloaded";

export interface ThreadRuntimeSnapshot {
  status: ThreadRuntimeStatus;
  threadId?: ThreadId;
  lastActiveAt?: number;
  unloadedAt?: number;
  idleUnloadMs: number;
}

// ============================================================
// Token 使用量
// ============================================================

export interface TokenUsage {
  inputTokens: number;
  outputTokens: number;
  cachedInputTokens?: number;
  reasoningOutputTokens?: number;
}

/** 合并 Token 使用量 */
export function mergeTokenUsage(a: TokenUsage, b: TokenUsage): TokenUsage {
  return {
    inputTokens: a.inputTokens + b.inputTokens,
    outputTokens: a.outputTokens + b.outputTokens,
    cachedInputTokens: (a.cachedInputTokens ?? 0) + (b.cachedInputTokens ?? 0),
    reasoningOutputTokens: (a.reasoningOutputTokens ?? 0) + (b.reasoningOutputTokens ?? 0),
  };
}

// ============================================================
// Compaction — 上下文压缩
// ============================================================

/** 压缩触发原因 */
export type CompactionReason =
  | "auto_token_limit"     // 自动：token 接近上限
  | "auto_pre_turn"        // 自动：新 Turn 开始前
  | "user_requested"       // 用户手动触发
  | "interrupted_resume"   // 中断后恢复时压缩
  | "model_changed"        // 模型切换后压缩，避免旧模型上下文污染新模型
  | "context_window_changed"; // 上下文窗口变更后压缩，适配新的窗口预算

/** 压缩策略配置 */
export interface CompactionConfig {
  /** 是否启用自动压缩 */
  enabled: boolean;
  /** 触发自动压缩的 token 阈值（默认 100000） */
  autoCompactTokenThreshold: number;
  /** mid-turn 压缩触发比例，默认 0.9 */
  midTurnThresholdRatio?: number;
  /** 压缩后保留的最近用户消息最大 token 数（默认 20000） */
  retainedUserMessageMaxTokens: number;
  /** 压缩后最多保留最近多少条用户消息；未设置时只按 token 预算控制 */
  retainedRecentItemCount?: number;
  /** 摘要生成失败后的重试次数，默认 2 */
  summaryRetryCount?: number;
  /** 摘要生成重试的首次等待时间，默认由 AgentLoop 控制 */
  summaryRetryBaseDelayMs?: number;
  /** 摘要生成重试的单次等待上限，默认由 AgentLoop 控制 */
  summaryRetryMaxDelayMs?: number;
  /** 摘要生成重试的退避倍率，默认由 AgentLoop 控制 */
  summaryRetryBackoffFactor?: number;
  /** rollout 超过该字节数后可生成 gzip 归档快照 */
  archiveRolloutAfterBytes?: number;
  /** 压缩提示词（可自定义） */
  compactPrompt?: string;
  /** 压缩摘要生成方式：local 使用当前模型，remote 调用远程压缩服务 */
  compactionProvider?: "local" | "remote";
  /** 远程压缩服务地址，要求兼容 { instruction, input, model? } 请求 */
  remoteCompactUrl?: string;
  /** 远程压缩服务鉴权 token，可选 */
  remoteCompactApiKey?: string;
  /** 远程压缩服务模型名，可选 */
  remoteCompactModel?: string;
  /** 模型上下文窗口大小（参考 Codex model_context_window），用于感知实际窗口上限 */
  contextWindowSize?: number;
}

export const DEFAULT_COMPACTION_CONFIG: CompactionConfig = {
  enabled: true,
  autoCompactTokenThreshold: 100_000,
  midTurnThresholdRatio: 0.9,
  retainedUserMessageMaxTokens: 20_000,
  summaryRetryCount: 2,
};

/** 上下文压缩启动参数，用于让客户端观察压缩流程 */
export interface ThreadCompactStartParams {
  reason: CompactionReason;
  itemCount: number;
  tokensBefore: number;
  tokenThreshold: number;
  contextWindowSize?: number;
  retryCount: number;
  timestamp: number;
}

// ============================================================
// Rollout — 会话持久化记录（参考 Codex RolloutItem）
// ============================================================

/** Rollout 记录行（JSONL 格式，每行一条） */
export interface RolloutLine {
  /** 时间戳 */
  timestamp: string;
  /** 记录内容 */
  item: RolloutItem;
}

/** Rollout 条目类型 */
export type RolloutItem =
  | { type: "session_meta"; meta: SessionMeta }
  | { type: "turn_item"; turnId: TurnId; item: TurnItem }
  | { type: "turn_usage"; turnId: TurnId; usage: TokenUsage }
  | {
      type: "compact_params";
      reason: CompactionReason;
      status: "started" | "completed" | "failed";
      itemCount: number;
      tokensBefore: number;
      tokensAfter?: number;
      error?: string;
    }
  | { type: "compacted"; summary: string; replacementHistory: TurnItem[] }
  | { type: "turn_context"; turnId: TurnId; cwd: string };

/** 会话元信息 */
export interface SessionMeta {
  id: ThreadId;
  timestamp: string;
  modelProvider: string;
  model?: string;
  /** 用户自定义名称；null 表示显式清除。 */
  name?: string | null;
  /** 所属文件夹路径（对应 pinnedFolders 的 path） */
  folderId?: string;
}

// ============================================================
// Agent 相关类型
// ============================================================

/** Agent 事件（参考 Codex EventMsg） */
type AgentEventThreadContext = { threadId?: ThreadId; clientId?: string };

export type AgentEvent =
  | ({ type: "turn_started"; turnId: TurnId } & AgentEventThreadContext)
  | ({ type: "turn_completed"; turnId: TurnId; usage?: TokenUsage } & AgentEventThreadContext)
  | ({ type: "turn_interrupted"; turnId: TurnId } & AgentEventThreadContext)
  | ({ type: "turn_failed"; turnId: TurnId; error: string } & AgentEventThreadContext)
  | ({ type: "item_started"; item: TurnItem } & AgentEventThreadContext)
  | ({ type: "item_completed"; item: TurnItem } & AgentEventThreadContext)
  | ({ type: "item_updated"; item: TurnItem } & AgentEventThreadContext)
  | ({ type: "thread_compact_started"; params: ThreadCompactStartParams } & AgentEventThreadContext)
  | ({ type: "context_compacted"; summary: string; tokensBefore: number; tokensAfter: number } & AgentEventThreadContext)
  | ({ type: "context_usage"; estimatedTokens: number; threshold: number; percentage: number; contextWindowSize: number } & AgentEventThreadContext)
  | ({ type: "stream_delta"; delta: string; itemType: "assistant_message" | "reasoning"; roundId?: number } & AgentEventThreadContext)
  | ({ type: "tool_approval_required"; toolCallId: string; toolName: string; arguments: Record<string, unknown>; riskLevel: ToolRiskLevel; description?: string } & AgentEventThreadContext)
  | ({ type: "error"; message: string } & AgentEventThreadContext)
  | ({ type: "warning"; message: string } & AgentEventThreadContext);

/** Agent Turn 输入 */
export interface AgentTurnInput {
  /** 用户输入内容 */
  content: string;
  /** 用户附加的文件列表 */
  attachments?: FileAttachment[];
  /** 客户端消息 ID */
  clientId?: string;
  /** 要执行的会话 ID；为空时创建新会话 */
  threadId?: string | null;
  /** 是否从中断恢复 */
  isResume?: boolean;
  /** 恢复时附带的上下文提示 */
  resumeContext?: string;
}

/** Agent Turn 回调（用于向渲染进程推送事件） */
export interface AgentTurnCallbacks {
  onEvent: (event: AgentEvent) => void;
  /**
   * 流式增量回调。
   * roundId 标识当前流式轮次（同一轮 tool 调用循环内 roundId 不变），
   * 跨轮切换时前端可主动清空 streamingReasoning/streamingContent，
   * 避免上一轮残留内容与本轮 delta 拼接（防止跨轮内容泄漏）。
   */
  onStreamDelta?: (
    delta: string,
    itemType: "assistant_message" | "reasoning",
    roundId?: number,
    threadId?: ThreadId,
    clientId?: string
  ) => void;
}

// ============================================================
// 工具定义
// ============================================================

/** 工具风险等级 */
export type ToolRiskLevel = "safe" | "moderate" | "dangerous";

/** 工具定义 */
export interface ToolDefinition {
  name: string;
  description: string;
  parameters: Record<string, unknown>; // JSON Schema
  riskLevel: ToolRiskLevel;
  /** 是否需要用户确认才能执行 */
  requiresApproval: boolean;
  /** 是否为文件删除类操作（range.clear / sheet.operation delete / ui.removeControl） */
  isFileDeletion?: boolean;
  /** 工具依赖的运行时环境。未设置或 "none" 表示无特殊依赖。 */
  requiresOfficeApp?: "excel" | "word" | "presentation" | "any";
}

/** 工具执行结果 */
export interface ToolExecutionResult {
  success: boolean;
  data?: unknown;
  error?: string;
}

/** 工具执行器接口 */
export interface ToolExecutor {
  readonly name: string;
  execute(args: Record<string, unknown>): Promise<ToolExecutionResult>;
}
