import type { createAIClient } from "../../providers/aiClient";
import type { AgentTurnCallbacks, Thread, Turn, TurnItem } from "../../shared/types";
import { mergeTokenUsage } from "../../shared/types";
import { collectStreamEvents, type StreamParams, type StreamResult } from "./streamCollector";
import { runAIRequestWithRetry, type AIRequestRetryConfig } from "./aiRequestRetry";

type AIClient = ReturnType<typeof createAIClient>;
type StreamResultWithErrorItem = StreamResult & { errorItem?: TurnItem };

export async function collectRoundStream(input: {
  aiClient: AIClient;
  streamParams: StreamParams;
  callbacks: AgentTurnCallbacks;
  round: number;
  retryConfig?: AIRequestRetryConfig;
  signal?: AbortSignal;
}): Promise<StreamResult> {
  let hasVisibleOutput = false;
  const callbacks: AgentTurnCallbacks = {
    ...input.callbacks,
    onEvent: (event) => {
      hasVisibleOutput = true;
      input.callbacks.onEvent(event);
    },
    onStreamDelta: input.callbacks.onStreamDelta
      ? (...args) => {
          hasVisibleOutput = true;
          input.callbacks.onStreamDelta?.(...args);
        }
      : undefined,
  };

  return runAIRequestWithRetry({
    phase: "sampling",
    config: input.retryConfig,
    signal: input.signal,
    canRetry: () => !hasVisibleOutput,
    operation: () =>
      collectStreamEvents(input.aiClient.streamChat(input.streamParams), callbacks, input.round),
  });
}

export async function emitStreamErrorItem(input: {
  streamResult: StreamResult;
  turn: Turn;
  callbacks: AgentTurnCallbacks;
  appendTurnItem: (threadId: string, turnId: string, item: TurnItem) => Promise<void>;
}): Promise<boolean> {
  const errorItem = (input.streamResult as StreamResultWithErrorItem).errorItem;
  if (!errorItem) return false;

  input.turn.items.push(errorItem);
  await input.appendTurnItem(input.turn.threadId, input.turn.turnId, errorItem);
  input.callbacks.onEvent({ type: "item_started", item: errorItem });
  input.callbacks.onEvent({ type: "item_completed", item: errorItem });
  input.callbacks.onEvent({
    type: "error",
    message: errorItem.type === "error" ? errorItem.message : "Unknown error",
  });
  return true;
}

export function applyStreamUsage(input: {
  streamResult: StreamResult;
  turn: Turn;
  activeThread: Thread | null;
}): void {
  if (!input.streamResult.usage) return;
  input.turn.tokenUsage = input.streamResult.usage;
  if (!input.activeThread) return;

  input.activeThread.metadata.totalTokenUsage = input.activeThread.metadata.totalTokenUsage
    ? mergeTokenUsage(input.activeThread.metadata.totalTokenUsage, input.streamResult.usage)
    : input.streamResult.usage;
}
