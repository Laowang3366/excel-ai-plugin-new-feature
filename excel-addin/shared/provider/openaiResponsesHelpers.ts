import type { AgentFinishReason, AgentStreamEvent, AgentTokenUsage } from "../agent/types";

export function extractResponsesErrorMessage(obj: Record<string, unknown>): string {
  // Official failed shape: { type, response: { status, error: { message } } }
  if (obj.response && typeof obj.response === "object") {
    const response = obj.response as Record<string, unknown>;
    if (typeof response.error === "string" && response.error) return response.error;
    if (response.error && typeof response.error === "object") {
      const msg = (response.error as { message?: unknown }).message;
      if (typeof msg === "string" && msg) return msg;
    }
  }
  if (typeof obj.error === "string" && obj.error) return obj.error;
  if (obj.error && typeof obj.error === "object") {
    const msg = (obj.error as { message?: unknown }).message;
    if (typeof msg === "string" && msg) return msg;
  }
  if (typeof obj.message === "string" && obj.message) return obj.message;
  return "";
}

export function parseResponsesUsage(raw: Record<string, unknown>): AgentTokenUsage | undefined {
  const input =
    typeof raw.input_tokens === "number"
      ? raw.input_tokens
      : typeof raw.prompt_tokens === "number"
        ? raw.prompt_tokens
        : undefined;
  const output =
    typeof raw.output_tokens === "number"
      ? raw.output_tokens
      : typeof raw.completion_tokens === "number"
        ? raw.completion_tokens
        : undefined;
  if (input == null || output == null) return undefined;
  const usage: AgentTokenUsage = { inputTokens: input, outputTokens: output };
  const inDetails = raw.input_tokens_details;
  if (inDetails && typeof inDetails === "object") {
    const cached = (inDetails as { cached_tokens?: unknown }).cached_tokens;
    if (typeof cached === "number") usage.cachedInputTokens = cached;
  }
  if (typeof raw.cached_tokens === "number") usage.cachedInputTokens = raw.cached_tokens;
  const outDetails = raw.output_tokens_details;
  if (outDetails && typeof outDetails === "object") {
    const reasoning = (outDetails as { reasoning_tokens?: unknown }).reasoning_tokens;
    if (typeof reasoning === "number") usage.reasoningOutputTokens = reasoning;
  }
  if (typeof raw.reasoning_tokens === "number") usage.reasoningOutputTokens = raw.reasoning_tokens;
  return usage;
}

export function mapIncompleteReason(reasonRaw: string): {
  reason: AgentFinishReason;
  rawReason: string;
} {
  let reason: AgentFinishReason = "unknown";
  if (reasonRaw === "max_output_tokens" || reasonRaw === "length") reason = "length";
  else if (reasonRaw === "content_filter") reason = "content_filter";
  return { reason, rawReason: reasonRaw };
}

export function finishEvent(
  reason: AgentFinishReason,
  rawReason?: string,
): AgentStreamEvent {
  return rawReason
    ? { type: "finish", reason, rawReason }
    : { type: "finish", reason };
}
