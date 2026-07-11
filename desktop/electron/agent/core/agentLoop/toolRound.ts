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
  await processCalls(
    input.streamResult.toolCalls,
    input.streamResult.pendingToolCallItems,
    input.turn,
    input.toolExecutors,
    input.approvalConfig,
    input.callbacks,
    input.appendTurnItem,
    input.appendToolExecutionLog,
    input.throwIfAborted
  );
  input.throwIfAborted();

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
