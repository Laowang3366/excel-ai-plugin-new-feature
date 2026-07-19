import type { AgentFinishReason, AgentStreamEvent, AgentTokenUsage } from "../agent/types";
import type { ToolNameMaps } from "./openaiToolNameMap";

interface ToolSlot {
  index: number;
  idParts: string[];
  nameParts: string[];
  pendingArgs: string[];
  args: string;
  began: boolean;
  ended: boolean;
  frozenId?: string;
  frozenExternalName?: string;
  frozenInternalName?: string;
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

          if (typeof item.id === "string" && item.id.length > 0) {
            const err = this.appendField(slot, "id", item.id);
            if (err) return err;
          }

          const fn = item.function;
          if (fn && typeof fn === "object") {
            const f = fn as Record<string, unknown>;
            if (typeof f.name === "string" && f.name.length > 0) {
              const err = this.appendField(slot, "name", f.name);
              if (err) return err;
            }
            if (typeof f.arguments === "string") {
              const begin = this.tryBegin(slot);
              if (begin && "error" in begin) return begin;
              if (begin && Array.isArray(begin)) events.push(...begin);

              if (!slot.began) {
                // Still waiting for complete id+name; buffer args in order.
                slot.pendingArgs.push(f.arguments);
              } else {
                slot.args += f.arguments;
                events.push({
                  type: "tool_call_delta",
                  toolCallId: slot.frozenId!,
                  argumentsDelta: f.arguments,
                });
              }
            }
          }
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
      slot = {
        index,
        idParts: [],
        nameParts: [],
        pendingArgs: [],
        args: "",
        began: false,
        ended: false,
      };
      this.slots.set(index, slot);
    }
    return slot;
  }

  private currentId(slot: ToolSlot): string {
    return slot.frozenId ?? slot.idParts.join("");
  }

  private currentName(slot: ToolSlot): string {
    return slot.frozenExternalName ?? slot.nameParts.join("");
  }

  /** Append id/name fragments; after freeze only identical full value is accepted. */
  private appendField(
    slot: ToolSlot,
    field: "id" | "name",
    piece: string,
  ): { error: string } | null {
    if (field === "id") {
      if (slot.frozenId != null) {
        if (piece === slot.frozenId) return null;
        return { error: `tool call id conflict at index ${slot.index}` };
      }
      slot.idParts.push(piece);
      return null;
    }
    if (slot.frozenExternalName != null) {
      if (piece === slot.frozenExternalName) return null;
      return { error: `tool call name conflict at index ${slot.index}` };
    }
    slot.nameParts.push(piece);
    return null;
  }

  /**
   * Begin only with non-empty id and exact external map key.
   * Prefix of a known key waits; unknown complete name is parse error.
   */
  private tryBegin(slot: ToolSlot): AgentStreamEvent[] | { error: string } | null {
    if (slot.began) return null;
    const id = this.currentId(slot);
    const external = this.currentName(slot);
    if (!id || !external) return null;

    const internal = this.maps.externalToInternal.get(external);
    if (internal == null) {
      for (const key of this.maps.externalToInternal.keys()) {
        if (key.startsWith(external) && key !== external) return null;
      }
      return { error: `unknown external tool name: ${external}` };
    }

    slot.began = true;
    slot.frozenId = id;
    slot.frozenExternalName = external;
    slot.frozenInternalName = internal;
    const events: AgentStreamEvent[] = [
      { type: "tool_call_begin", toolCallId: id, toolName: internal },
    ];
    for (const piece of slot.pendingArgs) {
      slot.args += piece;
      events.push({
        type: "tool_call_delta",
        toolCallId: id,
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
        const begin = this.tryBegin(slot);
        if (begin && "error" in begin) return begin;
        if (begin && Array.isArray(begin)) events.push(...begin);
      }

      if (!slot.began) {
        if (!this.currentId(slot)) return { error: "tool call missing id before finish" };
        if (!this.currentName(slot)) return { error: "tool call missing name before finish" };
        return { error: "tool call incomplete before finish" };
      }
      if (!slot.frozenId || !slot.frozenInternalName) {
        return { error: "tool call missing frozen id/name before end" };
      }

      // Include any args that arrived before begin but were not flushed (should be empty).
      for (const piece of slot.pendingArgs) slot.args += piece;
      slot.pendingArgs = [];

      events.push({
        type: "tool_call_end",
        toolCallId: slot.frozenId,
        toolName: slot.frozenInternalName,
        argumentsJson: slot.args === "" ? "{}" : slot.args,
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
