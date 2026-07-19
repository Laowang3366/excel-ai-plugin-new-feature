import type { AgentStreamEvent } from "../agent/types";
import type { ToolNameMaps } from "./openaiToolNameMap";
import {
  extractResponsesErrorMessage,
  finishEvent,
  mapIncompleteReason,
  parseResponsesUsage,
} from "./openaiResponsesHelpers";

interface TextSlot {
  emitted: string;
}

interface ToolSlot {
  itemId: string;
  callId?: string;
  externalName?: string;
  internalName?: string;
  deltaArgs: string;
  /** Present when function_call_arguments.done provided arguments (including ""). */
  doneArgs?: string;
  hasDoneArgs: boolean;
  /** Present when output_item.done provided arguments (including ""). */
  itemDoneArgs?: string;
  hasItemDoneArgs: boolean;
  itemDone: boolean;
  argsDone: boolean;
  began: boolean;
  ended: boolean;
}

type IngestResult = AgentStreamEvent[] | { error: string; provider?: boolean };

/**
 * OpenAI Responses SSE → AgentStreamEvent.
 * Tool slots keyed by item_id; agent toolCallId is always frozen call_id.
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
          error: extractResponsesErrorMessage(obj) || "provider error event",
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
    const full =
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
      slot = {
        itemId,
        deltaArgs: "",
        hasDoneArgs: false,
        hasItemDoneArgs: false,
        itemDone: false,
        argsDone: false,
        began: false,
        ended: false,
      };
      this.tools.set(itemId, slot);
      this.hasAnyTool = true;
    }
    return slot;
  }

  /** Stable setter: only identical full call_id/name repeats are allowed. */
  private setCallId(slot: ToolSlot, value: string): { error: string } | null {
    if (!value) return null;
    if (slot.callId != null && slot.callId !== value) {
      return { error: `tool call_id conflict at item ${slot.itemId}` };
    }
    slot.callId = value;
    return null;
  }

  private setExternalName(slot: ToolSlot, value: string): { error: string } | null {
    if (!value) return null;
    if (slot.externalName != null && slot.externalName !== value) {
      return { error: `tool name conflict at item ${slot.itemId}` };
    }
    slot.externalName = value;
    return null;
  }

  private applyIdentity(
    slot: ToolSlot,
    src: Record<string, unknown>,
  ): { error: string } | null {
    if (typeof src.call_id === "string" && src.call_id) {
      const err = this.setCallId(slot, src.call_id);
      if (err) return err;
    }
    if (typeof src.name === "string" && src.name) {
      const err = this.setExternalName(slot, src.name);
      if (err) return err;
    }
    return null;
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
    const idErr = this.applyIdentity(slot, it);
    if (idErr) return idErr;

    // output_item.added may carry draft arguments; not a final marker.
    if (!done && typeof it.arguments === "string" && !slot.hasItemDoneArgs && !slot.hasDoneArgs) {
      // Keep as provisional delta only if no final args yet and no deltas.
      if (!slot.deltaArgs && it.arguments) slot.deltaArgs = it.arguments;
    }

    const events: AgentStreamEvent[] = [];
    const begin = this.tryBegin(slot);
    if (begin && "error" in begin) return begin;
    if (Array.isArray(begin)) events.push(...begin);

    if (done) {
      slot.itemDone = true;
      if (typeof it.arguments === "string") {
        slot.itemDoneArgs = it.arguments; // may be ""
        slot.hasItemDoneArgs = true;
      }
      const end = this.endSlot(slot, /*requireFinalMarker*/ false);
      if (end && "error" in end) return end;
      if (Array.isArray(end)) events.push(...end);
    }
    return events;
  }

  private onArgsDelta(obj: Record<string, unknown>): IngestResult {
    const itemId = typeof obj.item_id === "string" ? obj.item_id : "";
    if (!itemId) return { error: "function_call_arguments.delta missing item_id" };
    const slot = this.ensure(itemId);
    const idErr = this.applyIdentity(slot, obj);
    if (idErr) return idErr;
    const delta = typeof obj.delta === "string" ? obj.delta : "";
    if (!delta) return [];

    const events: AgentStreamEvent[] = [];
    if (!slot.began) {
      slot.deltaArgs += delta;
      const begin = this.tryBegin(slot);
      if (begin && "error" in begin) return begin;
      if (Array.isArray(begin)) events.push(...begin);
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
    const idErr = this.applyIdentity(slot, obj);
    if (idErr) return idErr;
    if (typeof obj.arguments === "string") {
      slot.doneArgs = obj.arguments; // may be ""
      slot.hasDoneArgs = true;
    }
    slot.argsDone = true;
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
    // Freeze identity for all subsequent events.
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

  /**
   * Priority: item.done.arguments (incl "") > args.done.arguments (incl "") > delta > "{}".
   * Presence of key/field is tracked via has* flags so empty string wins over deltas.
   */
  private finalArgs(slot: ToolSlot): string {
    if (slot.hasItemDoneArgs) return slot.itemDoneArgs === "" ? "{}" : (slot.itemDoneArgs ?? "{}");
    if (slot.hasDoneArgs) return slot.doneArgs === "" ? "{}" : (slot.doneArgs ?? "{}");
    if (slot.deltaArgs !== "") return slot.deltaArgs;
    return "{}";
  }

  /**
   * @param requireFinalMarker when true (terminal flush), need itemDone or argsDone.
   */
  private endSlot(
    slot: ToolSlot,
    requireFinalMarker: boolean,
  ): AgentStreamEvent[] | { error: string } | null {
    if (slot.ended) return null;
    if (requireFinalMarker && !slot.itemDone && !slot.argsDone) {
      return {
        error: `tool slot ${slot.itemId} incomplete at terminal (need item.done or arguments.done)`,
      };
    }

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
      const part = this.endSlot(slot, /*requireFinalMarker*/ true);
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
        const usage = parseResponsesUsage(src.usage as Record<string, unknown>);
        if (usage) events.push({ type: "usage", usage });
      }
    }
    const flushed = this.flushAllTools();
    if ("error" in flushed) return flushed;
    events.push(...flushed);

    if (kind === "completed") {
      this.pendingFinish = finishEvent(this.hasAnyTool ? "tool_calls" : "stop");
    } else {
      const reasonRaw =
        typeof response.incomplete_details === "object" &&
        response.incomplete_details &&
        typeof (response.incomplete_details as { reason?: unknown }).reason === "string"
          ? String((response.incomplete_details as { reason: string }).reason)
          : typeof obj.reason === "string"
            ? obj.reason
            : "incomplete";
      const mapped = mapIncompleteReason(reasonRaw);
      this.pendingFinish = finishEvent(mapped.reason, mapped.rawReason);
    }
    this.sawTerminal = true;
    return events;
  }
}
