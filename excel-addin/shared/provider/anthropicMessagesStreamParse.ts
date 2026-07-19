import type {
  AgentFinishReason,
  AgentStreamEvent,
  AgentTokenUsage,
} from "../agent/types";
import type { ToolNameMaps } from "./openaiToolNameMap";

type BlockKind = "text" | "tool" | "ignored";

interface BlockSlot {
  index: number;
  kind: BlockKind;
  stopped: boolean;
  callId?: string;
  internalName?: string;
  externalName?: string;
  args: string;
  began: boolean;
  ended: boolean;
}

type IngestResult = AgentStreamEvent[] | { error: string; provider?: boolean };

/**
 * Anthropic Messages SSE → AgentStreamEvent.
 * Tool slots keyed by content_block index; toolCallId is real tool_use id.
 */
export class AnthropicMessagesStreamAssembler {
  private blocks = new Map<number, BlockSlot>();
  private usage: AgentTokenUsage = { inputTokens: 0, outputTokens: 0 };
  private knownInput = false;
  private knownOutput = false;
  private stopReason: AgentFinishReason | undefined;
  private rawStopReason: string | undefined;
  private emittedFinish = false;

  constructor(private readonly maps: ToolNameMaps) {}

  ingest(data: unknown): IngestResult {
    if (data == null || typeof data !== "object") {
      return { error: "SSE data is not an object" };
    }
    const obj = data as Record<string, unknown>;
    const type = typeof obj.type === "string" ? obj.type : "";
    if (!type) return [];

    switch (type) {
      case "message_start":
        return this.onMessageStart(obj);
      case "content_block_start":
        return this.onBlockStart(obj);
      case "content_block_delta":
        return this.onBlockDelta(obj);
      case "content_block_stop":
        return this.onBlockStop(obj);
      case "message_delta":
        return this.onMessageDelta(obj);
      case "message_stop":
        // Finish is emitted only from finalize() on [DONE]/EOF.
        return [];
      case "error":
        return {
          error: extractErrorMessage(obj) || "provider error event",
          provider: true,
        };
      case "ping":
      default:
        return [];
    }
  }

  finalize(): AgentStreamEvent[] | { error: string } {
    if (this.stopReason == null) {
      return { error: "stream ended without stop_reason" };
    }
    for (const slot of this.blocks.values()) {
      if (slot.kind === "tool" && !slot.ended) {
        return { error: `tool content block ${slot.index} not stopped before end` };
      }
    }
    if (this.emittedFinish) return [];
    this.emittedFinish = true;
    if (
      this.rawStopReason &&
      (this.stopReason === "unknown" || this.stopReason === "content_filter")
    ) {
      return [
        {
          type: "finish",
          reason: this.stopReason,
          rawReason: this.rawStopReason,
        },
      ];
    }
    return [{ type: "finish", reason: this.stopReason }];
  }

  private emitUsage(): AgentStreamEvent {
    return {
      type: "usage",
      usage: {
        inputTokens: this.usage.inputTokens,
        outputTokens: this.usage.outputTokens,
        ...(this.usage.cachedInputTokens != null
          ? { cachedInputTokens: this.usage.cachedInputTokens }
          : {}),
        ...(this.usage.reasoningOutputTokens != null
          ? { reasoningOutputTokens: this.usage.reasoningOutputTokens }
          : {}),
      },
    };
  }

  private mergeUsage(raw: Record<string, unknown>): void {
    // Field-level cumulative state; do not double-add.
    if (typeof raw.input_tokens === "number") {
      this.usage.inputTokens = raw.input_tokens;
      this.knownInput = true;
    }
    if (typeof raw.output_tokens === "number") {
      this.usage.outputTokens = raw.output_tokens;
      this.knownOutput = true;
    }
    if (typeof raw.cache_read_input_tokens === "number") {
      this.usage.cachedInputTokens = raw.cache_read_input_tokens;
    }
  }

  /** Emit only when both input and output tokens are known (complete snapshot). */
  private maybeEmitUsage(): AgentStreamEvent[] {
    if (!this.knownInput || !this.knownOutput) return [];
    return [this.emitUsage()];
  }

  private onMessageStart(obj: Record<string, unknown>): IngestResult {
    const message = obj.message;
    if (message && typeof message === "object") {
      const usage = (message as { usage?: unknown }).usage;
      if (usage && typeof usage === "object") {
        this.mergeUsage(usage as Record<string, unknown>);
        return this.maybeEmitUsage();
      }
    }
    return [];
  }

  private onBlockStart(obj: Record<string, unknown>): IngestResult {
    if (typeof obj.index !== "number" || !Number.isInteger(obj.index)) {
      return { error: "content_block_start missing integer index" };
    }
    const index = obj.index;
    if (this.blocks.has(index)) {
      return { error: `duplicate content_block_start index ${index}` };
    }
    const block = obj.content_block;
    if (!block || typeof block !== "object") {
      return { error: "content_block_start missing content_block" };
    }
    const b = block as Record<string, unknown>;
    const bType = typeof b.type === "string" ? b.type : "";

    if (bType === "text") {
      this.blocks.set(index, {
        index,
        kind: "text",
        stopped: false,
        args: "",
        began: false,
        ended: false,
      });
      const initial = typeof b.text === "string" ? b.text : "";
      return initial ? [{ type: "text_delta", delta: initial }] : [];
    }

    if (bType === "tool_use") {
      const id = typeof b.id === "string" ? b.id : "";
      const name = typeof b.name === "string" ? b.name : "";
      if (!id) return { error: "tool_use block missing id" };
      if (!name) return { error: "tool_use block missing name" };
      const internal = this.maps.externalToInternal.get(name);
      if (internal == null) {
        return { error: `unknown external tool name: ${name}` };
      }
      this.blocks.set(index, {
        index,
        kind: "tool",
        stopped: false,
        callId: id,
        externalName: name,
        internalName: internal,
        args: "",
        began: true,
        ended: false,
      });
      return [
        {
          type: "tool_call_begin",
          toolCallId: id,
          toolName: internal,
        },
      ];
    }

    this.blocks.set(index, {
      index,
      kind: "ignored",
      stopped: false,
      args: "",
      began: false,
      ended: false,
    });
    return [];
  }

  private onBlockDelta(obj: Record<string, unknown>): IngestResult {
    if (typeof obj.index !== "number" || !Number.isInteger(obj.index)) {
      return { error: "content_block_delta missing integer index" };
    }
    const slot = this.blocks.get(obj.index);
    if (!slot) return { error: `content_block_delta for unknown index ${obj.index}` };
    if (slot.stopped) return { error: `content_block_delta after stop at index ${obj.index}` };

    const delta = obj.delta;
    if (!delta || typeof delta !== "object") {
      return { error: "content_block_delta missing delta" };
    }
    const d = delta as Record<string, unknown>;
    const dType = typeof d.type === "string" ? d.type : "";

    if (dType === "text_delta") {
      if (slot.kind !== "text") {
        return { error: `text_delta on non-text block ${obj.index}` };
      }
      const text = typeof d.text === "string" ? d.text : "";
      return text ? [{ type: "text_delta", delta: text }] : [];
    }

    if (dType === "input_json_delta") {
      if (slot.kind !== "tool" || !slot.callId) {
        return { error: `input_json_delta on non-tool block ${obj.index}` };
      }
      const partial = typeof d.partial_json === "string" ? d.partial_json : "";
      if (!partial) return [];
      slot.args += partial;
      return [
        {
          type: "tool_call_delta",
          toolCallId: slot.callId,
          argumentsDelta: partial,
        },
      ];
    }

    if (dType === "thinking_delta" || dType === "signature_delta") {
      return [];
    }

    if (slot.kind === "ignored") return [];
    return { error: `unsupported content_block_delta type: ${dType || "(missing)"}` };
  }

  private onBlockStop(obj: Record<string, unknown>): IngestResult {
    if (typeof obj.index !== "number" || !Number.isInteger(obj.index)) {
      return { error: "content_block_stop missing integer index" };
    }
    const slot = this.blocks.get(obj.index);
    if (!slot) return { error: `content_block_stop for unknown index ${obj.index}` };
    if (slot.stopped) return [];
    slot.stopped = true;

    if (slot.kind === "tool") {
      if (!slot.callId || !slot.internalName) {
        return { error: `tool block ${obj.index} missing id/name at stop` };
      }
      if (slot.ended) return [];
      slot.ended = true;
      return [
        {
          type: "tool_call_end",
          toolCallId: slot.callId,
          toolName: slot.internalName,
          argumentsJson: slot.args === "" ? "{}" : slot.args,
        },
      ];
    }
    return [];
  }

  private onMessageDelta(obj: Record<string, unknown>): IngestResult {
    const events: AgentStreamEvent[] = [];
    const usage = obj.usage;
    if (usage && typeof usage === "object") {
      this.mergeUsage(usage as Record<string, unknown>);
      events.push(...this.maybeEmitUsage());
    }
    const delta = obj.delta;
    if (delta && typeof delta === "object") {
      const stop = (delta as { stop_reason?: unknown }).stop_reason;
      if (typeof stop === "string" && stop) {
        const mapped = mapStopReason(stop);
        this.stopReason = mapped.reason;
        this.rawStopReason = mapped.rawReason;
      }
    }
    if (typeof obj.stop_reason === "string" && obj.stop_reason) {
      const mapped = mapStopReason(obj.stop_reason);
      this.stopReason = mapped.reason;
      this.rawStopReason = mapped.rawReason;
    }
    return events;
  }
}

function mapStopReason(stop: string): {
  reason: AgentFinishReason;
  rawReason?: string;
} {
  switch (stop) {
    case "end_turn":
      return { reason: "stop" };
    case "tool_use":
      return { reason: "tool_calls" };
    case "max_tokens":
    case "model_context_window_exceeded":
      return { reason: "length" };
    case "stop_sequence":
      return { reason: "stop" };
    case "refusal":
      return { reason: "content_filter", rawReason: stop };
    default:
      return { reason: "unknown", rawReason: stop };
  }
}

function extractErrorMessage(obj: Record<string, unknown>): string {
  if (typeof obj.error === "string") return obj.error;
  if (obj.error && typeof obj.error === "object") {
    const err = obj.error as { message?: unknown; type?: unknown };
    if (typeof err.message === "string") return err.message;
  }
  // Nested Anthropic error shapes: { error: { error: { message } } } unlikely;
  // also support top-level message.
  if (typeof obj.message === "string") return obj.message;
  return "";
}
