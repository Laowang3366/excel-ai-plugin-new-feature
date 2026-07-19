import type {
  AgentFinishReason,
  AgentStreamEvent,
  AgentTokenUsage,
} from "../agent/types";
import type { ToolNameMaps } from "./openaiToolNameMap";

interface TextSlot {
  emitted: string;
}

interface ToolSlot {
  itemId: string;
  callId?: string;
  externalName?: string;
  internalName?: string;
  deltaArgs: string;
  doneArgs?: string;
  itemDoneArgs?: string;
  began: boolean;
  ended: boolean;
}

type IngestResult = AgentStreamEvent[] | { error: string; provider?: boolean };

/**
 * OpenAI Responses SSE → AgentStreamEvent.
 * Tool slots keyed by item_id; agent toolCallId is always real call_id.
 */
export class OpenAiResponsesStreamAssembler {
  private textSlots = new Map<string, TextSlot>();
  private tools = new Map<string, ToolSlot>();
  private sawTerminal = false;
  private pendingFinish: AgentStreamEvent | undefined;
  private emittedFinish = false;
  private hasAnyTool = false;

  constructor(private readonly maps: ToolNameMaps) {}

  ingest(data: unknown): IngestResult {
    if (data == null || typeof data !== "object") {
      return { error: "SSE data is not an object" };
    }
    const obj = data as Record<string, unknown>;
    const type = typeof obj.type === "string" ? obj.type : "";
    if (!type) return [];

    switch (type) {
      case "response.output_text.delta":
        return this.onTextDelta(obj);
      case "response.output_text.done":
      case "response.content_part.done":
        return this.onTextDone(obj);
      case "response.output_item.added":
        return this.onItem(obj, false);
      case "response.output_item.done":
        return this.onItem(obj, true);
      case "response.function_call_arguments.delta":
        return this.onArgsDelta(obj);
      case "response.function_call_arguments.done":
        return this.onArgsDone(obj);
      case "response.completed":
        return this.onTerminal(obj, "completed");
      case "response.incomplete":
        return this.onTerminal(obj, "incomplete");
      case "response.failed":
      case "error":
        return {
          error: extractErrorMessage(obj) || "provider error event",
          provider: true,
        };
      default:
        return [];
    }
  }

  finalize(): AgentStreamEvent[] | { error: string } {
    if (!this.sawTerminal || !this.pendingFinish) {
      return { error: "stream ended without response.completed/incomplete" };
    }
    if (this.emittedFinish) return [];
    this.emittedFinish = true;
    return [this.pendingFinish];
  }

  private textKey(obj: Record<string, unknown>): string {
    const itemId = typeof obj.item_id === "string" ? obj.item_id : "";
    const index =
      typeof obj.content_index === "number"
        ? obj.content_index
        : typeof obj.output_index === "number"
          ? obj.output_index
          : 0;
    return `${itemId}:${index}`;
  }

  private onTextDelta(obj: Record<string, unknown>): IngestResult {
    const delta = typeof obj.delta === "string" ? obj.delta : "";
    if (!delta) return [];
    const key = this.textKey(obj);
    const slot = this.textSlots.get(key) ?? { emitted: "" };
    slot.emitted += delta;
    this.textSlots.set(key, slot);
    return [{ type: "text_delta", delta }];
  }

  private onTextDone(obj: Record<string, unknown>): IngestResult {
    let full =
      typeof obj.text === "string"
        ? obj.text
        : typeof (obj.part as { text?: unknown } | undefined)?.text === "string"
          ? String((obj.part as { text: string }).text)
          : "";
    if (!full) return [];
    const key = this.textKey(obj);
    const slot = this.textSlots.get(key) ?? { emitted: "" };
    if (!full.startsWith(slot.emitted)) return [];
    const suffix = full.slice(slot.emitted.length);
    slot.emitted = full;
    this.textSlots.set(key, slot);
    return suffix ? [{ type: "text_delta", delta: suffix }] : [];
  }

  private ensure(itemId: string): ToolSlot {
    let slot = this.tools.get(itemId);
    if (!slot) {
      slot = { itemId, deltaArgs: "", began: false, ended: false };
      this.tools.set(itemId, slot);
      this.hasAnyTool = true;
    }
    return slot;
  }

  private applyFunctionFields(slot: ToolSlot, src: Record<string, unknown>): void {
    if (typeof src.call_id === "string" && src.call_id) slot.callId = src.call_id;
    if (typeof src.name === "string" && src.name) slot.externalName = src.name;
    if (typeof src.arguments === "string") slot.itemDoneArgs = src.arguments;
  }

  private onItem(obj: Record<string, unknown>, done: boolean): IngestResult {
    const item = obj.item;
    if (!item || typeof item !== "object") return [];
    const it = item as Record<string, unknown>;
    if (it.type !== "function_call") return [];
    const itemId =
      typeof it.id === "string" && it.id
        ? it.id
        : typeof obj.item_id === "string"
          ? obj.item_id
          : "";
    if (!itemId) return { error: "function_call item missing item id" };
    const slot = this.ensure(itemId);
    this.applyFunctionFields(slot, it);
    const events: AgentStreamEvent[] = [];
    const begin = this.tryBegin(slot);
    if (begin && "error" in begin) return begin;
    if (Array.isArray(begin)) events.push(...begin);
    if (done) {
      const end = this.endSlot(slot);
      if (end && "error" in end) return end;
      if (Array.isArray(end)) events.push(...end);
    }
    return events;
  }

  private onArgsDelta(obj: Record<string, unknown>): IngestResult {
    const itemId = typeof obj.item_id === "string" ? obj.item_id : "";
    if (!itemId) return { error: "function_call_arguments.delta missing item_id" };
    const slot = this.ensure(itemId);
    const delta = typeof obj.delta === "string" ? obj.delta : "";
    if (!delta) return [];
    const events: AgentStreamEvent[] = [];
    if (!slot.began) {
      slot.deltaArgs += delta;
      const begin = this.tryBegin(slot);
      if (begin && "error" in begin) return begin;
      if (Array.isArray(begin)) {
        events.push(...begin);
        // tryBegin already replays full deltaArgs once.
      }
    } else {
      slot.deltaArgs += delta;
      events.push({
        type: "tool_call_delta",
        toolCallId: slot.callId!,
        argumentsDelta: delta,
      });
    }
    return events;
  }

  private onArgsDone(obj: Record<string, unknown>): IngestResult {
    const itemId = typeof obj.item_id === "string" ? obj.item_id : "";
    if (!itemId) return { error: "function_call_arguments.done missing item_id" };
    const slot = this.ensure(itemId);
    if (typeof obj.arguments === "string") slot.doneArgs = obj.arguments;
    if (typeof obj.call_id === "string" && obj.call_id) slot.callId = obj.call_id;
    if (typeof obj.name === "string" && obj.name) slot.externalName = obj.name;
    const begin = this.tryBegin(slot);
    if (begin && "error" in begin) return begin;
    return Array.isArray(begin) ? begin : [];
  }

  private tryBegin(slot: ToolSlot): AgentStreamEvent[] | { error: string } | null {
    if (slot.began) return null;
    if (!slot.callId || !slot.externalName) return null;
    const internal = this.maps.externalToInternal.get(slot.externalName);
    if (internal == null) {
      for (const key of this.maps.externalToInternal.keys()) {
        if (key.startsWith(slot.externalName) && key !== slot.externalName) return null;
      }
      return { error: `unknown external tool name: ${slot.externalName}` };
    }
    slot.began = true;
    slot.internalName = internal;
    const events: AgentStreamEvent[] = [
      { type: "tool_call_begin", toolCallId: slot.callId, toolName: internal },
    ];
    if (slot.deltaArgs) {
      events.push({
        type: "tool_call_delta",
        toolCallId: slot.callId,
        argumentsDelta: slot.deltaArgs,
      });
    }
    return events;
  }

  private finalArgs(slot: ToolSlot): string {
    if (typeof slot.itemDoneArgs === "string" && slot.itemDoneArgs !== "") {
      return slot.itemDoneArgs;
    }
    if (typeof slot.doneArgs === "string" && slot.doneArgs !== "") return slot.doneArgs;
    if (slot.deltaArgs !== "") return slot.deltaArgs;
    return "{}";
  }

  /** Begin (if needed) + end; returns events or error. */
  private endSlot(slot: ToolSlot): AgentStreamEvent[] | { error: string } | null {
    if (slot.ended) return null;
    const events: AgentStreamEvent[] = [];
    if (!slot.began) {
      const begin = this.tryBegin(slot);
      if (begin && "error" in begin) return begin;
      if (Array.isArray(begin)) events.push(...begin);
    }
    if (!slot.began) {
      if (!slot.callId) return { error: "tool call missing call_id before end" };
      if (!slot.externalName) return { error: "tool call missing name before end" };
      return { error: "tool call incomplete before end" };
    }
    if (!slot.callId || !slot.internalName) {
      return { error: "tool call missing call_id/name before end" };
    }
    slot.ended = true;
    events.push({
      type: "tool_call_end",
      toolCallId: slot.callId,
      toolName: slot.internalName,
      argumentsJson: this.finalArgs(slot),
    });
    return events;
  }

  private flushAllTools(): AgentStreamEvent[] | { error: string } {
    const events: AgentStreamEvent[] = [];
    for (const slot of this.tools.values()) {
      const part = this.endSlot(slot);
      if (part && "error" in part) return part;
      if (Array.isArray(part)) events.push(...part);
    }
    return events;
  }

  private onTerminal(
    obj: Record<string, unknown>,
    kind: "completed" | "incomplete",
  ): IngestResult {
    const events: AgentStreamEvent[] = [];
    const response = (
      obj.response && typeof obj.response === "object" ? obj.response : obj
    ) as Record<string, unknown>;
    for (const src of [response, obj]) {
      if (src.usage && typeof src.usage === "object") {
        const usage = parseUsage(src.usage as Record<string, unknown>);
        if (usage) events.push({ type: "usage", usage });
      }
    }
    const flushed = this.flushAllTools();
    if ("error" in flushed) return flushed;
    events.push(...flushed);

    if (kind === "completed") {
      this.pendingFinish = {
        type: "finish",
        reason: this.hasAnyTool ? "tool_calls" : "stop",
      };
    } else {
      const reasonRaw =
        typeof response.incomplete_details === "object" &&
        response.incomplete_details &&
        typeof (response.incomplete_details as { reason?: unknown }).reason === "string"
          ? String((response.incomplete_details as { reason: string }).reason)
          : typeof obj.reason === "string"
            ? obj.reason
            : "incomplete";
      let reason: AgentFinishReason = "unknown";
      if (reasonRaw === "max_output_tokens" || reasonRaw === "length") reason = "length";
      else if (reasonRaw === "content_filter") reason = "content_filter";
      this.pendingFinish = { type: "finish", reason, rawReason: reasonRaw };
    }
    this.sawTerminal = true;
    return events;
  }
}

function extractErrorMessage(obj: Record<string, unknown>): string {
  if (typeof obj.error === "string") return obj.error;
  if (obj.error && typeof obj.error === "object") {
    const msg = (obj.error as { message?: unknown }).message;
    if (typeof msg === "string") return msg;
  }
  if (typeof obj.message === "string") return obj.message;
  return "";
}

function parseUsage(raw: Record<string, unknown>): AgentTokenUsage | undefined {
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
