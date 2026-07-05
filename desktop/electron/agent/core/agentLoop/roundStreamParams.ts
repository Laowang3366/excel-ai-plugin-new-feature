import type { StateRuntimeStore } from "../../memory/stateRuntimeStore";
import type { AgentTurnInput, ToolExecutor, TurnItem } from "../../shared/types";
import type { AIClientConfig, ReasoningMode } from "../../providers/aiClient";
import { turnItemGroupsToChatMessages } from "../../shared/messageBuilder";
import { resolveImageAttachments } from "../../attachments/imageAttachmentResolver";
import {
  appendRuntimeLongTermMemoryContext,
  buildEffectiveSystemPrompt,
} from "./buildStreamParams";
import { resolveMaxTokens } from "./maxTokens";
import {
  getToolDefinitions,
} from "./toolExecutor";
import type { StreamParams } from "./streamCollector";

export async function buildRoundStreamParams(input: {
  turnItemGroups: TurnItem[][];
  turnInput: AgentTurnInput;
  aiConfig: AIClientConfig;
  configuredReasoningMode?: ReasoningMode;
  baseSystemPrompt?: string;
  folderId?: string;
  stateRuntimeStore?: StateRuntimeStore;
  toolExecutors?: Map<string, ToolExecutor>;
  signal?: AbortSignal;
  round: number;
  resumeContext?: string;
}): Promise<{
  streamParams: StreamParams;
  effectiveSystemPrompt: string;
  toolDefs: ReturnType<typeof getToolDefinitions>;
}> {
  const messages = turnItemGroupsToChatMessages(input.turnItemGroups);
  await resolveImageAttachments(messages);

  if (input.resumeContext) {
    messages.push({
      role: "system",
      content: input.resumeContext,
    });
  }

  let effectiveSystemPrompt = await buildEffectiveSystemPrompt(
    input.baseSystemPrompt,
    input.folderId,
    {
      content: input.turnInput.content,
      attachments: input.turnInput.attachments,
    }
  );
  effectiveSystemPrompt = await appendRuntimeLongTermMemoryContext(
    effectiveSystemPrompt,
    input.stateRuntimeStore
  );

  const toolDefs = getToolDefinitions(input.toolExecutors);
  return {
    effectiveSystemPrompt,
    toolDefs,
    streamParams: {
      messages,
      tools: toolDefs,
      systemPrompt: effectiveSystemPrompt,
      maxTokens: resolveMaxTokens(input.aiConfig),
      reasoningMode: input.aiConfig.reasoningMode || input.configuredReasoningMode || "high",
      signal: input.signal,
      roundId: input.round,
    },
  };
}
