import { abortableDelay, throwIfAborted } from "./streamProvider";
import type {
  AgentStreamEvent,
  AgentStreamProvider,
  StreamChatRequest,
} from "./types";

export type ScriptedRound =
  | AgentStreamEvent[]
  | ((ctx: {
      callCount: number;
      request: StreamChatRequest;
    }) => AgentStreamEvent[] | Promise<AgentStreamEvent[]>);

export interface ScriptedStreamProviderOptions {
  rounds: ScriptedRound[];
  onExhausted?: "stop" | "error";
  eventDelayMs?: number;
}

export class ScriptedStreamProvider implements AgentStreamProvider {
  readonly rounds: ScriptedRound[];
  readonly onExhausted: "stop" | "error";
  readonly eventDelayMs: number;
  callCount = 0;
  lastRequest: StreamChatRequest | undefined;

  constructor(options: ScriptedStreamProviderOptions) {
    this.rounds = options.rounds;
    this.onExhausted = options.onExhausted ?? "error";
    this.eventDelayMs = options.eventDelayMs ?? 0;
  }

  async *streamChat(request: StreamChatRequest): AsyncIterable<AgentStreamEvent> {
    throwIfAborted(request.signal);
    this.callCount += 1;
    this.lastRequest = {
      systemPrompt: request.systemPrompt,
      messages: request.messages.slice(),
      tools: request.tools.slice(),
      signal: request.signal,
    };

    const index = this.callCount - 1;
    if (index >= this.rounds.length) {
      if (this.onExhausted === "stop") {
        yield { type: "finish", reason: "stop" };
        return;
      }
      yield {
        type: "error",
        message: "Scripted provider exhausted",
        kind: "provider",
      };
      return;
    }

    const script = this.rounds[index];
    const events =
      typeof script === "function"
        ? await script({ callCount: this.callCount, request: this.lastRequest })
        : script;

    for (const event of events) {
      throwIfAborted(request.signal);
      if (this.eventDelayMs > 0) {
        await abortableDelay(this.eventDelayMs, request.signal);
      }
      throwIfAborted(request.signal);
      yield event;
    }
  }
}

export function textThenStop(text: string): AgentStreamEvent[] {
  return [
    { type: "text_delta", delta: text },
    { type: "finish", reason: "stop" },
  ];
}

export function toolCallThenFinish(
  toolCallId: string,
  toolName: string,
  argumentsJson: string,
  finish: "tool_calls" | "stop" = "tool_calls",
): AgentStreamEvent[] {
  return [
    { type: "tool_call_begin", toolCallId, toolName },
    { type: "tool_call_delta", toolCallId, argumentsDelta: argumentsJson },
    {
      type: "tool_call_end",
      toolCallId,
      toolName,
      argumentsJson,
    },
    { type: "finish", reason: finish },
  ];
}

export function errorEvent(
  message: string,
  kind: "provider" | "http" | "network" | "aborted" = "provider",
  status?: number,
): AgentStreamEvent[] {
  return [{ type: "error", message, kind, status }];
}
