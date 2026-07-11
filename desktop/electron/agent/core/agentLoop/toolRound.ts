import type {
  AgentTurnCallbacks,
  CompactionConfig,
  ToolDefinition,
  ToolExecutor,
  Turn,
  TurnItem,
} from "../../shared/types";
import { estimateRequestTokens } from "../../memory/compaction";
import { turnItemGroupsToChatMessages } from "../../shared/messageBuilder";
import type { StreamResult } from "./streamCollector";
import {
  processToolCalls,
  type ToolApprovalConfig,
  type ToolExecutionLogRecord,
} from "./toolExecutor";
import {
  buildFormulaKnowledgeQuery,
  buildFormulaSceneQuery,
  guardFormulaWorkflowExecutors,
  shouldLoadFormulaMethodology,
  shouldSearchFormulaScene,
} from "./formulaKnowledgePolicy";
import {
  buildPendingFormulaValidationRead,
  shouldRunFormulaVerification,
} from "./formulaVerification";
import { createFormulaWorkflowTurnView } from "./formulaWorkflowContext";

type ProcessToolCalls = typeof processToolCalls;

export async function handleToolRound(input: {
  streamResult: StreamResult;
  turn: Turn;
  toolExecutors: Map<string, ToolExecutor>;
  approvalConfig: ToolApprovalConfig;
  callbacks: AgentTurnCallbacks;
  appendTurnItem: (threadId: string, turnId: string, item: TurnItem) => Promise<void>;
  appendToolExecutionLog?: (record: ToolExecutionLogRecord) => Promise<void>;
  turnItemGroups: TurnItem[][];
  effectiveSystemPrompt: string;
  toolDefs: ToolDefinition[];
  compactionConfig: CompactionConfig;
  runMidTurnCompaction: () => Promise<void>;
  throwIfAborted: () => void;
  processToolCallsImpl?: ProcessToolCalls;
}): Promise<boolean> {
  if (input.streamResult.toolCalls.length === 0) return false;

  const processCalls = input.processToolCallsImpl ?? processToolCalls;
  const workflowTurn = createFormulaWorkflowTurnView(input.turn, input.turnItemGroups);
  const guardedExecutors = guardFormulaWorkflowExecutors(input.toolExecutors, workflowTurn);
  await processCalls(
    input.streamResult.toolCalls,
    input.streamResult.pendingToolCallItems,
    input.turn,
    guardedExecutors,
    input.approvalConfig,
    input.callbacks,
    input.appendTurnItem,
    input.appendToolExecutionLog,
    input.throwIfAborted
  );
  input.throwIfAborted();

  if (shouldLoadFormulaMethodology(workflowTurn, input.toolExecutors)) {
    const toolCallId = `auto-formula-knowledge-${input.turn.turnId}-${Date.now()}`;
    await processCalls(
      [{
        id: toolCallId,
        name: "knowledge.search",
        arguments: JSON.stringify({
          query: buildFormulaKnowledgeQuery(workflowTurn),
          topK: 3,
          scope: "formula_methodology",
        }),
      }],
      new Map(),
      input.turn,
      input.toolExecutors,
      input.approvalConfig,
      input.callbacks,
      input.appendTurnItem,
      input.appendToolExecutionLog,
      input.throwIfAborted,
    );
    input.throwIfAborted();
  }

  if (shouldSearchFormulaScene(workflowTurn, input.toolExecutors)) {
    const toolCallId = `auto-formula-scene-${input.turn.turnId}-${Date.now()}`;
    await processCalls(
      [{
        id: toolCallId,
        name: "knowledge.search",
        arguments: JSON.stringify({
          query: buildFormulaSceneQuery(workflowTurn),
          topK: 3,
          scope: "formula_scene",
        }),
      }],
      new Map(),
      input.turn,
      guardedExecutors,
      input.approvalConfig,
      input.callbacks,
      input.appendTurnItem,
      input.appendToolExecutionLog,
      input.throwIfAborted,
    );
    input.throwIfAborted();
  }

  const validationRead = buildPendingFormulaValidationRead(workflowTurn);
  if (validationRead) {
    const toolCallId = `auto-formula-read-${input.turn.turnId}-${Date.now()}`;
    await processCalls(
      [{
        id: toolCallId,
        name: "range.read",
        arguments: JSON.stringify(validationRead),
      }],
      new Map(),
      input.turn,
      guardedExecutors,
      input.approvalConfig,
      input.callbacks,
      input.appendTurnItem,
      input.appendToolExecutionLog,
      input.throwIfAborted,
    );
    input.throwIfAborted();
  }

  if (shouldRunFormulaVerification(workflowTurn)) {
    const toolCallId = `auto-formula-verify-${input.turn.turnId}-${Date.now()}`;
    await processCalls(
      [{ id: toolCallId, name: "formula.verify", arguments: "{}" }],
      new Map(),
      input.turn,
      guardedExecutors,
      input.approvalConfig,
      input.callbacks,
      input.appendTurnItem,
      input.appendToolExecutionLog,
      input.throwIfAborted,
    );
    input.throwIfAborted();
  }

  if (shouldRunMidTurnCompaction({
    turnItemGroups: input.turnItemGroups,
    systemPrompt: input.effectiveSystemPrompt,
    tools: input.toolDefs,
    compactionConfig: input.compactionConfig,
  })) {
    await input.runMidTurnCompaction();
  }
  input.throwIfAborted();
  return true;
}

export function shouldRunMidTurnCompaction(input: {
  turnItemGroups: TurnItem[][];
  systemPrompt: string;
  tools: ToolDefinition[];
  compactionConfig: CompactionConfig;
}): boolean {
  const allTokens = estimateRequestTokens({
    messages: turnItemGroupsToChatMessages(input.turnItemGroups),
    systemPrompt: input.systemPrompt,
    tools: input.tools,
  });
  const midTurnRatio = input.compactionConfig.midTurnThresholdRatio ?? 0.9;
  return input.compactionConfig.enabled
    && allTokens > input.compactionConfig.autoCompactTokenThreshold * midTurnRatio;
}
