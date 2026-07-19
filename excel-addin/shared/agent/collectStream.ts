import { throwIfAborted } from "./streamProvider";
import type {
  AgentFinishReason,
  AgentRoundStreamResult,
  AgentStreamEvent,
  AgentStreamError,
  AgentTokenUsage,
  AgentToolCall,
} from "./types";

interface Slot {
  id: string;
  name: string;
  delta: string;
  endArgs?: string;
  ended: boolean;
}

export interface CollectAgentStreamOptions {
  signal?: AbortSignal;
  onTextDelta?: (delta: string) => void;
}

function emptyUsage(): AgentTokenUsage {
  return { inputTokens: 0, outputTokens: 0 };
}

function asOptions(
  signalOrOptions?: AbortSignal | CollectAgentStreamOptions,
): CollectAgentStreamOptions {
  if (signalOrOptions == null) return {};
  if (typeof signalOrOptions === "object" && "aborted" in signalOrOptions) {
    return { signal: signalOrOptions as AbortSignal };
  }
  return signalOrOptions as CollectAgentStreamOptions;
}

/** Aggregate one stream; only ended tool calls, first-seen order; no JSON.parse. */
export async function collectAgentStream(
  stream: AsyncIterable<AgentStreamEvent>,
  signalOrOptions?: AbortSignal | CollectAgentStreamOptions,
): Promise<AgentRoundStreamResult> {
  const { signal, onTextDelta } = asOptions(signalOrOptions);
  throwIfAborted(signal);
  let assistantText = "";
  const order: string[] = [];
  const slots = new Map<string, Slot>();
  let usage: AgentTokenUsage | undefined;
  let finishReason: AgentFinishReason | undefined;
  let error: AgentStreamError | undefined;

  const ensure = (id: string): Slot => {
    let slot = slots.get(id);
    if (!slot) {
      slot = { id, name: "", delta: "", ended: false };
      slots.set(id, slot);
      order.push(id);
    }
    return slot;
  };

  for await (const event of stream) {
    throwIfAborted(signal);
    switch (event.type) {
      case "text_delta":
        assistantText += event.delta;
        onTextDelta?.(event.delta);
        break;
      case "tool_call_begin": {
        const slot = ensure(event.toolCallId);
        if (event.toolName && !slot.name) slot.name = event.toolName;
        break;
      }
      case "tool_call_delta":
        ensure(event.toolCallId).delta += event.argumentsDelta;
        break;
      case "tool_call_end": {
        const slot = ensure(event.toolCallId);
        if (slot.ended) break;
        if (event.toolName) slot.name = event.toolName;
        if (typeof event.argumentsJson === "string" && event.argumentsJson !== "") {
          slot.endArgs = event.argumentsJson;
        }
        slot.ended = true;
        break;
      }
      case "usage":
        usage = { ...event.usage };
        break;
      case "finish":
        finishReason = event.reason;
        break;
      case "error":
        error = {
          message: event.message,
          kind: event.kind,
          status: event.status,
          url: event.url,
        };
        break;
      default:
        break;
    }
    if (error) break;
  }

  throwIfAborted(signal);

  const toolCalls: AgentToolCall[] = [];
  for (const id of order) {
    const slot = slots.get(id);
    if (!slot || !slot.ended) continue;
    let argumentsJson = "{}";
    if (typeof slot.endArgs === "string" && slot.endArgs !== "") {
      argumentsJson = slot.endArgs;
    } else if (slot.delta !== "") {
      argumentsJson = slot.delta;
    }
    toolCalls.push({ id: slot.id, name: slot.name, argumentsJson });
  }

  return {
    assistantText,
    toolCalls,
    finishReason: finishReason ?? "stop",
    usage,
    error,
  };
}

export function sumUsage(total: AgentTokenUsage, next?: AgentTokenUsage): AgentTokenUsage {
  if (!next) return total;
  const out: AgentTokenUsage = {
    inputTokens: total.inputTokens + next.inputTokens,
    outputTokens: total.outputTokens + next.outputTokens,
  };
  if (total.cachedInputTokens != null || next.cachedInputTokens != null) {
    out.cachedInputTokens = (total.cachedInputTokens ?? 0) + (next.cachedInputTokens ?? 0);
  }
  if (total.reasoningOutputTokens != null || next.reasoningOutputTokens != null) {
    out.reasoningOutputTokens =
      (total.reasoningOutputTokens ?? 0) + (next.reasoningOutputTokens ?? 0);
  }
  return out;
}

export { emptyUsage };
