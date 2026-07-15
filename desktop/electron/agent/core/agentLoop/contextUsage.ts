import {
  DEFAULT_COMPACTION_CONFIG,
  type CompactionConfig,
  type Thread,
  type ToolDefinition,
  type Turn,
  type TurnItem,
} from "../../shared/types";
import { estimateRequestTokens } from "../../memory/compaction";
import { turnItemGroupsToChatMessages } from "../../shared/messageBuilder";

export interface PromptHistoryState {
  activeThread: Thread | null;
  activeTurn: Turn | null;
  compactedHistory: TurnItem[] | null;
}

export function collectPromptTurnItems(state: PromptHistoryState): TurnItem[] {
  if (!state.activeThread) return [];

  const items = state.compactedHistory ? [...state.compactedHistory] : [];
  for (const turn of state.activeThread.turns) {
    items.push(...turn.items);
  }
  if (state.activeTurn) {
    items.push(...state.activeTurn.items);
  }
  return items;
}

export function collectPromptTurnItemGroups(state: PromptHistoryState): TurnItem[][] {
  if (!state.activeThread) return [];

  const groups: TurnItem[][] = [];
  if (state.compactedHistory) {
    groups.push(state.compactedHistory);
  }
  for (const turn of state.activeThread.turns) {
    groups.push(turn.items);
  }
  if (state.activeTurn) {
    groups.push(state.activeTurn.items);
  }
  return groups.filter((items) => items.length > 0);
}

export function buildContextUsageEvent(input: {
  groups: TurnItem[][];
  activeThread: Thread | null;
  compactionConfig?: CompactionConfig;
  systemPrompt?: string;
  tools: ToolDefinition[];
}) {
  const config = input.compactionConfig ?? DEFAULT_COMPACTION_CONFIG;
  const contextWindowSize =
    input.activeThread?.metadata?.contextWindowSize || config.contextWindowSize || 128_000;
  const estimatedTokens = estimateRequestTokens({
    messages: turnItemGroupsToChatMessages(input.groups),
    systemPrompt: input.systemPrompt,
    tools: input.tools,
  });
  return {
    type: "context_usage" as const,
    estimatedTokens,
    threshold: config.autoCompactTokenThreshold,
    percentage: Math.min(Math.round((estimatedTokens / contextWindowSize) * 100), 100),
    contextWindowSize,
  };
}
