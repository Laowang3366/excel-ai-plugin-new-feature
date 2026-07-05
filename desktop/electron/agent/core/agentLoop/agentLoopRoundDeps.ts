import type { StateRuntimeStore } from "../../memory/stateRuntimeStore";
import type {
  AgentTurnCallbacks,
  AgentTurnInput,
  CompactionConfig,
  Thread,
  ThreadId,
  ToolDefinition,
  ToolExecutor,
  Turn,
  TurnItem,
} from "../../shared/types";
import type { AIClientConfig, ReasoningMode } from "../../providers/aiClient";
import type { createAIClient } from "../../providers/aiClient";
import type { AIRequestRetryConfig } from "./aiRequestRetry";
import type { ToolExecutionLogRecord } from "./toolExecutor";
import { runAgentLoopRounds } from "./agentLoopRunner";

type AIClient = ReturnType<typeof createAIClient>;

export async function runAgentLoopWithDeps(input: {
  turn: Turn;
  callbacks: AgentTurnCallbacks;
  turnInput: AgentTurnInput;
  resumeContext?: string;
  aiClient: AIClient;
  aiConfig: AIClientConfig;
  configuredReasoningMode?: ReasoningMode;
  baseSystemPrompt?: string;
  folderId?: string;
  stateRuntimeStore?: StateRuntimeStore;
  toolExecutors?: Map<string, ToolExecutor>;
  permissionMode?: "normal" | "auto_approve_safe" | "confirm_all";
  requestToolApproval?: (params: {
    toolCallId: string;
    toolName: string;
    arguments: Record<string, unknown>;
    riskLevel: "safe" | "moderate" | "dangerous";
    description?: string;
  }) => Promise<{ approved: boolean; alwaysAllow?: boolean }>;
  samplingRetryConfig?: AIRequestRetryConfig;
  signal?: AbortSignal;
  appendTurnItem: (threadId: ThreadId, turnId: string, item: TurnItem) => Promise<void>;
  appendToolExecutionLog?: (record: ToolExecutionLogRecord) => Promise<void>;
  getTurnItemGroups: () => TurnItem[][];
  getActiveThread: () => Thread | null;
  getSessionCompactionConfig: () => CompactionConfig;
  runMidTurnCompaction: () => Promise<void>;
  emitContextUsage: (
    callbacks: AgentTurnCallbacks,
    requestContext: { systemPrompt: string; tools: ToolDefinition[] }
  ) => void;
  throwIfAborted: () => void;
}): Promise<void> {
  await runAgentLoopRounds({
    ...input,
    toolExecutors: input.toolExecutors!,
    approvalConfig: {
      permissionMode: input.permissionMode || "normal",
      requestToolApproval: input.requestToolApproval,
    },
  });
}
