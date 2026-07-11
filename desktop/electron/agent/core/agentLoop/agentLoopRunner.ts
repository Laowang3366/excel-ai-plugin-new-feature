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
import { buildRoundStreamParams } from "./roundStreamParams";
import { emitStreamResultItems as emitCollectedStreamResultItems } from "./streamResultItems";
import {
  applyStreamUsage,
  collectRoundStream,
  emitStreamErrorItem,
} from "./streamRound";
import { handleToolRound } from "./toolRound";
import type {
  ToolApprovalConfig,
  ToolExecutionLogRecord,
} from "./toolExecutor";
import {
  formatFormulaDiagnostic,
  getFormulaCompletionDiagnostic,
} from "./formulaKnowledgePolicy";
import { createToolResultItem } from "./toolResultItems";
import { createFormulaWorkflowTurnView } from "./formulaWorkflowContext";

type AIClient = ReturnType<typeof createAIClient>;

export interface AgentLoopRunnerInput {
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
  toolExecutors: Map<string, ToolExecutor>;
  approvalConfig: ToolApprovalConfig;
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
}

export async function runAgentLoopRounds(input: AgentLoopRunnerInput): Promise<void> {
  let round = 0;
  let formulaCompletionRetryCount = 0;

  while (true) {
    input.throwIfAborted();
    round++;

    const { streamParams, effectiveSystemPrompt, toolDefs } = await buildRoundStreamParams({
      turnItemGroups: input.getTurnItemGroups(),
      turnInput: input.turnInput,
      aiConfig: input.aiConfig,
      configuredReasoningMode: input.configuredReasoningMode,
      baseSystemPrompt: input.baseSystemPrompt,
      folderId: input.folderId,
      stateRuntimeStore: input.stateRuntimeStore,
      toolExecutors: input.toolExecutors,
      signal: input.signal,
      round,
      resumeContext: input.resumeContext,
    });

    const streamResult = await collectRoundStream({
      aiClient: input.aiClient,
      streamParams,
      callbacks: input.callbacks,
      round,
      retryConfig: input.samplingRetryConfig,
      signal: input.signal,
    });

    if (await emitStreamErrorItem({
      streamResult,
      turn: input.turn,
      callbacks: input.callbacks,
      appendTurnItem: input.appendTurnItem,
    })) {
      return;
    }

    input.throwIfAborted();
    await emitCollectedStreamResultItems({
      streamResult,
      turn: input.turn,
      callbacks: input.callbacks,
      appendTurnItem: input.appendTurnItem,
    });
    input.throwIfAborted();

    if (await handleToolRound({
      streamResult,
      turn: input.turn,
      toolExecutors: input.toolExecutors,
      approvalConfig: input.approvalConfig,
      callbacks: input.callbacks,
      appendTurnItem: input.appendTurnItem,
      appendToolExecutionLog: input.appendToolExecutionLog,
      turnItemGroups: input.getTurnItemGroups(),
      effectiveSystemPrompt,
      toolDefs,
      compactionConfig: input.getSessionCompactionConfig(),
      runMidTurnCompaction: input.runMidTurnCompaction,
      throwIfAborted: input.throwIfAborted,
    })) {
      formulaCompletionRetryCount = 0;
      continue;
    }

    const workflowTurn = createFormulaWorkflowTurnView(input.turn, input.getTurnItemGroups());
    const formulaDiagnostic = getFormulaCompletionDiagnostic(workflowTurn);
    if (formulaDiagnostic) {
      if (!formulaDiagnostic.retryable || formulaCompletionRetryCount >= 2) {
        throw new Error(formatFormulaDiagnostic(formulaDiagnostic));
      }
      formulaCompletionRetryCount += 1;
      await appendFormulaCompletionFeedback({
        turn: input.turn,
        callbacks: input.callbacks,
        appendTurnItem: input.appendTurnItem,
        diagnostic: formatFormulaDiagnostic(formulaDiagnostic),
        retryCount: formulaCompletionRetryCount,
      });
      continue;
    }

    applyStreamUsage({
      streamResult,
      turn: input.turn,
      activeThread: input.getActiveThread(),
    });
    input.emitContextUsage(input.callbacks, {
      systemPrompt: effectiveSystemPrompt,
      tools: toolDefs,
    });
    break;
  }
}

async function appendFormulaCompletionFeedback(input: {
  turn: Turn;
  callbacks: AgentTurnCallbacks;
  appendTurnItem: (threadId: string, turnId: string, item: TurnItem) => Promise<void>;
  diagnostic: string;
  retryCount: number;
}): Promise<void> {
  const toolCallId = `formula-completion-gate-${input.turn.turnId}-${input.retryCount}`;
  const callItem: TurnItem = {
    type: "tool_call",
    id: toolCallId,
    toolName: "formula.verify",
    arguments: { completionGate: true, retryCount: input.retryCount },
    status: "failed",
    timestamp: Date.now(),
  };
  const resultItem = createToolResultItem({
    toolCallId,
    toolName: "formula.verify",
    result: input.diagnostic,
    isError: true,
  });
  input.turn.items.push(callItem, resultItem);
  await input.appendTurnItem(input.turn.threadId, input.turn.turnId, callItem);
  await input.appendTurnItem(input.turn.threadId, input.turn.turnId, resultItem);
  input.callbacks.onEvent({ type: "item_started", item: callItem });
  input.callbacks.onEvent({ type: "item_completed", item: callItem });
  input.callbacks.onEvent({ type: "item_started", item: resultItem });
  input.callbacks.onEvent({ type: "item_completed", item: resultItem });
}
