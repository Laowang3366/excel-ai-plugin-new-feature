import type { AgentFinishReason, AgentStreamEvent, AgentTokenUsage } from "../agent/types";
import type { ToolNameMaps } from "./openaiToolNameMap";

interface ToolSlot {
  index: number;
  id?: string;
  externalName?: string;
  args: string;
  began: boolean;
  ended: boolean;
  pendingArgs: string[];
}

export class OpenAiChatStreamAssembler {
  private slots = new Map<number, ToolSlot>();
  private sawFinishReason = false;
  private pendingFinish: AgentStreamEvent | undefined;
  private emittedFinish = false;

  constructor(private readonly maps: ToolNameMaps) {}

  get hasFinishReason(): boolean {
    return this.sawFinishReason;
  }

  get finished(): boolean {
    return this.emittedFinish;
  }

  /** Parse one SSE data JSON object into events, or return parse error. */
  ingest(data: unknown): AgentStreamEvent[] | { error: string } {
    if (data == null || typeof data !== "object") {
      return { error: "SSE data is not an object" };
    }
    const obj = data as Record<string, unknown>;
    const events: AgentStreamEvent[] = [];

    if (obj.usage && typeof obj.usage === "object") {
      const usage = parseUsage(obj.usage as Record<string, unknown>);
      if (usage) events.push({ type: "usage", usage });
    }

    const choices = obj.choices;
    if (Array.isArray(choices) && choices.length > 0) {
      const choice = choices[0] as Record<string, unknown>;
      const delta = (choice.delta ?? {}) as Record<string, unknown>;

      if (typeof delta.content === "string" && delta.content.length > 0) {
        events.push({ type: "text_delta", delta: delta.content });
      }

      if (Array.isArray(delta.tool_calls)) {
        for (const raw of delta.tool_calls) {
          if (!raw || typeof raw !== "object") {
            return { error: "tool_calls item is not an object" };
          }
          const item = raw as Record<string, unknown>;
          if (typeof item.index !== "number" || !Number.isInteger(item.index)) {
            return { error: "tool_calls item missing integer index" };
          }
          const slot = this.ensure(item.index);
          if (typeof item.id === "string" && item.id) slot.id = item.id;
          const fn = item.function;
          if (fn && typeof fn === "object") {
            const f = fn as Record<string, unknown>;
            if (typeof f.name === "string" && f.name) slot.externalName = f.name;
            if (typeof f.arguments === "string") {
              if (!slot.began) slot.pendingArgs.push(f.arguments);
              else {
                slot.args += f.arguments;
                events.push({
                  type: "tool_call_delta",
                  toolCallId: slot.id!,
                  argumentsDelta: f.arguments,
                });
              }
            }
          }
          const begin = this.tryBegin(slot);
          if (begin && "error" in begin) return begin;
          if (begin && Array.isArray(begin)) events.push(...begin);
        }
      }

      const finishReason = choice.finish_reason;
      if (finishReason != null && finishReason !== "") {
        if (typeof finishReason !== "string") {
          return { error: "finish_reason must be a string" };
        }
        if (!this.sawFinishReason) {
          const endEvents = this.endAllOpen();
          if ("error" in endEvents) return endEvents;
          events.push(...endEvents);
          this.pendingFinish = mapFinish(finishReason);
          this.sawFinishReason = true;
        }
      }
    }

    return events;
  }

  /** Emit buffered finish as the last normal event when stream completes. */
  finalize(): AgentStreamEvent[] | { error: string } {
    if (!this.sawFinishReason || !this.pendingFinish) {
      return { error: "stream ended without finish_reason" };
    }
    if (this.emittedFinish) return [];
    this.emittedFinish = true;
    return [this.pendingFinish];
  }

  private ensure(index: number): ToolSlot {
    let slot = this.slots.get(index);
    if (!slot) {
      slot = { index, args: "", began: false, ended: false, pendingArgs: [] };
      this.slots.set(index, slot);
    }
    return slot;
  }

  private tryBegin(slot: ToolSlot): AgentStreamEvent[] | { error: string } | null {
    if (slot.began) return null;
    if (!slot.id || !slot.externalName) return null;
    const internal = this.maps.externalToInternal.get(slot.externalName);
    if (internal == null) {
      return { error: `unknown external tool name: ${slot.externalName}` };
    }
    slot.began = true;
    const events: AgentStreamEvent[] = [
      { type: "tool_call_begin", toolCallId: slot.id, toolName: internal },
    ];
    for (const piece of slot.pendingArgs) {
      slot.args += piece;
      events.push({
        type: "tool_call_delta",
        toolCallId: slot.id,
        argumentsDelta: piece,
      });
    }
    slot.pendingArgs = [];
    return events;
  }

  private endAllOpen(): AgentStreamEvent[] | { error: string } {
    const events: AgentStreamEvent[] = [];
    const ordered = [...this.slots.values()].sort((a, b) => a.index - b.index);
    for (const slot of ordered) {
      if (slot.ended) continue;
      if (!slot.began) {
        if (slot.id || slot.externalName || slot.pendingArgs.length || slot.args) {
          if (!slot.id) return { error: "tool call missing id before finish" };
          if (!slot.externalName) return { error: "tool call missing name before finish" };
          return { error: "tool call incomplete before finish" };
        }
        continue;
      }
      if (!slot.id) return { error: "tool call missing id before end" };
      const argumentsJson = slot.args === "" ? "{}" : slot.args;
      const toolName = this.maps.externalToInternal.get(slot.externalName ?? "");
      events.push({
        type: "tool_call_end",
        toolCallId: slot.id,
        toolName,
        argumentsJson,
      });
      slot.ended = true;
    }
    return events;
  }
}

function mapFinish(reason: string): AgentStreamEvent {
  const known: AgentFinishReason[] = ["stop", "tool_calls", "length", "content_filter"];
  if ((known as string[]).includes(reason)) {
    return { type: "finish", reason: reason as AgentFinishReason };
  }
  return { type: "finish", reason: "unknown", rawReason: reason };
}

function parseUsage(raw: Record<string, unknown>): AgentTokenUsage | undefined {
  const input =
    typeof raw.prompt_tokens === "number"
      ? raw.prompt_tokens
      : typeof raw.input_tokens === "number"
        ? raw.input_tokens
        : undefined;
  const output =
    typeof raw.completion_tokens === "number"
      ? raw.completion_tokens
      : typeof raw.output_tokens === "number"
        ? raw.output_tokens
        : undefined;
  if (input == null || output == null) return undefined;
  const usage: AgentTokenUsage = { inputTokens: input, outputTokens: output };
  const details = raw.prompt_tokens_details;
  if (details && typeof details === "object") {
    const cached = (details as { cached_tokens?: unknown }).cached_tokens;
    if (typeof cached === "number") usage.cachedInputTokens = cached;
  }
  if (typeof raw.cached_tokens === "number") usage.cachedInputTokens = raw.cached_tokens;
  const outDetails = raw.completion_tokens_details;
  if (outDetails && typeof outDetails === "object") {
    const reasoning = (outDetails as { reasoning_tokens?: unknown }).reasoning_tokens;
    if (typeof reasoning === "number") usage.reasoningOutputTokens = reasoning;
  }
  if (typeof raw.reasoning_tokens === "number") usage.reasoningOutputTokens = raw.reasoning_tokens;
  return usage;
}
