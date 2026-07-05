import type { AgentTurnCallbacks, Turn, TurnItem } from "../../shared/types";
import { collectStreamEvents } from "./streamCollector";

type StreamResult = Awaited<ReturnType<typeof collectStreamEvents>>;

export async function emitStreamResultItems(input: {
  streamResult: StreamResult;
  turn: Turn;
  callbacks: AgentTurnCallbacks;
  appendTurnItem: (threadId: string, turnId: string, item: TurnItem) => Promise<void>;
}): Promise<void> {
  const { streamResult, turn, callbacks, appendTurnItem } = input;

  if (streamResult.reasoningContent.length > 0 || streamResult.reasoningSummary.length > 0) {
    const reasoningItem: TurnItem = {
      type: "reasoning",
      id: `reasoning-${Date.now()}`,
      summaryText: streamResult.reasoningSummary,
      rawContent: streamResult.reasoningContent,
      timestamp: Date.now(),
    };
    turn.items.push(reasoningItem);
    await appendTurnItem(turn.threadId, turn.turnId, reasoningItem);
    callbacks.onEvent({ type: "item_started", item: reasoningItem });
    callbacks.onEvent({ type: "item_completed", item: reasoningItem });
  }

  if (streamResult.assistantContent) {
    const msgItem: TurnItem = {
      type: "assistant_message",
      id: `msg-${Date.now()}`,
      content: streamResult.assistantContent,
      phase: streamResult.toolCalls.length > 0 ? "commentary" : "final",
      timestamp: Date.now(),
    };
    turn.items.push(msgItem);
    await appendTurnItem(turn.threadId, turn.turnId, msgItem);
    callbacks.onEvent({ type: "item_started", item: msgItem });
    callbacks.onEvent({ type: "item_completed", item: msgItem });
  }

  for (const toolCall of streamResult.toolCalls) {
    const existingItem = streamResult.pendingToolCallItems.get(toolCall.id);
    if (!existingItem) continue;
    turn.items.push(existingItem);
    await appendTurnItem(turn.threadId, turn.turnId, existingItem);
    callbacks.onEvent({ type: "item_started", item: existingItem });
  }
}
