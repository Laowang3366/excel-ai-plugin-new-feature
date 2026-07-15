import type {
  AgentEvent,
  FileAttachment,
  ThreadMetadata,
  TokenUsage,
  TurnItem,
} from "../electronApi";
import type { StreamDeltaInput } from "./chatStreamBuffer";

export interface ChatState {
  /** 当前 Turn 的消息列表（只从 Agent 事件产出） */
  messages: TurnItem[];
  /** 是否正在流式输出 */
  isStreaming: boolean;
  /** 流式输出的增量内容（助手消息，仅用于实时展示） */
  streamingContent: string;
  /** 流式输出的推理增量（仅用于实时展示） */
  streamingReasoning: string;
  /**
   * 当前流式轮次 ID。
   * 当新 round 的 streamDelta 到达时，与上一轮不同就会清空
   * streamingContent/streamingReasoning，防止跨轮内容泄漏。
   */
  activeStreamingRound: number | null;
  /** 是否显示推理过程 */
  showReasoning: boolean;
  /** 推理面板的展开状态（按消息 ID） */
  reasoningExpanded: Record<string, boolean>;
  /** 当前 Turn ID */
  activeTurnId: string | null;
  /** 当前活跃的会话 ID */
  activeThreadId: string | null;
  /** 当前前端发起但尚未绑定 threadId 的请求 ID */
  activeClientId: string | null;
  /** 正在执行中的会话 ID 映射，用于切换会话后仍展示运行态 */
  runningThreadIds: Record<string, boolean>;
  /** 已发送停止请求、等待主进程确认的会话 ID 映射 */
  pendingInterruptThreadIds: Record<string, boolean>;
  /** 用户明确点击停止的会话 ID，用于屏蔽旧快照里的 in_progress 状态 */
  stoppedThreadIds: Record<string, boolean>;
  /** Turn 状态 */
  turnStatus: "idle" | "in_progress" | "completed" | "interrupted" | "failed";
  /** 上次中断的上下文提示 */
  lastInterruptContext: string | null;
  /** Token 使用量 */
  tokenUsage: TokenUsage | null;
  /** 上下文使用情况 */
  contextUsage: {
    estimatedTokens: number;
    threshold: number;
    percentage: number;
    contextWindowSize: number;
  } | null;
  /** 压缩提示 */
  compactionNotice: string | null;
  /** 错误信息 */
  error: string | null;

  /** 所有会话列表 */
  threads: ThreadMetadata[];

  /** 挂起等待审批的工具调用 */
  pendingToolCall: {
    id: string;
    toolName: string;
    arguments: Record<string, unknown>;
    riskLevel: "safe" | "moderate" | "dangerous";
    description?: string;
    canAlwaysAllow?: boolean;
  } | null;

  /** 待添加到输入框的文件列表 */
  pendingComposerFiles: FileAttachment[];
  /** 新建会话时暂存的文件夹 ID */
  pendingFolderId: string | null;
}

export interface ChatActions {
  sendMessage: (content: string, attachments?: FileAttachment[]) => Promise<string | null>;
  resumeFromInterruption: (content: string, attachments?: FileAttachment[]) => Promise<void>;
  interruptTurn: () => Promise<void>;
  toggleReasoning: (itemId?: string) => void;
  clearError: () => void;
  clearMessages: () => void;
  handleAgentEvent: (event: AgentEvent) => void;
  handleStreamDelta: (data: StreamDeltaInput) => void;
  loadThreads: () => Promise<void>;
  switchThread: (threadId: string) => Promise<void>;
  createNewThread: (folderId?: string) => Promise<void>;
  deleteThread: (threadId: string) => Promise<void>;
  moveThreadToFolder: (threadId: string, folderId?: string) => Promise<void>;
  confirmToolCall: (toolCallId: string, alwaysAllow?: boolean) => void;
  cancelToolCall: (toolCallId: string) => void;
  addFilesToComposer: (files: FileAttachment[]) => void;
  consumePendingFiles: () => FileAttachment[];
}
