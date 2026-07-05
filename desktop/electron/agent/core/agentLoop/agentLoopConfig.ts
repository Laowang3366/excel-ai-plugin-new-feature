import type { LongTermMemoryStore } from "../../memory/longTerm/memoryStore";
import type {
  CompactionConfig,
  ToolExecutor,
} from "../../shared/types";
import type {
  AIClientConfig,
  ReasoningMode,
} from "../../providers/aiClient";
import type { AIRequestPhase, AIRequestRetryConfig } from "./aiRequestRetry";
import type { CompactionProvider } from "./compactionProvider";

export interface AgentLoopConfig {
  /** AI 客户端配置 */
  aiConfig: AIClientConfig;
  /** 系统提示词 */
  systemPrompt?: string;
  /** 压缩配置 */
  compactionConfig?: CompactionConfig;
  /** 压缩摘要生成器；默认使用当前 AI 客户端本地生成 */
  compactionProvider?: CompactionProvider;
  /** 工具执行器映射 */
  toolExecutors?: Map<string, ToolExecutor>;
  /** 权限模式 */
  permissionMode?: "normal" | "auto_approve_safe" | "confirm_all";
  /** 推理力度（覆盖 aiConfig 中的 reasoningMode */
  reasoningMode?: ReasoningMode;
  /** 空闲多久后卸载内存中的 activeThread，默认 30 分钟；<=0 表示禁用 */
  threadIdleUnloadMs?: number;
  /** 请求工具审批的回调（由主进程提供，向渲染进程发送审批请求并等待响应） */
  requestToolApproval?: (params: {
    toolCallId: string;
    toolName: string;
    arguments: Record<string, unknown>;
    riskLevel: "safe" | "moderate" | "dangerous";
    description?: string;
  }) => Promise<{ approved: boolean; alwaysAllow?: boolean }>;
  /** 模型请求重试配置。sampling 用于普通回复，compact 用于上下文压缩。 */
  aiRequestRetryConfig?: Partial<Record<AIRequestPhase, AIRequestRetryConfig>>;
  /** Turn 完成后自动抽取并写入长期记忆。 */
  memoryStore?: LongTermMemoryStore;
}
