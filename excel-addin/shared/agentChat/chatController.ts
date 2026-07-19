import { AgentLoop } from "../agent/agentLoop";
import type { AgentMessage, AgentRunResult, LoopEvent } from "../agent/types";
import type { HostAdapter } from "../host/hostAdapter";
import { createStreamProviderFromStore } from "../provider/createStreamProvider";
import type { ProviderFetch } from "../provider/client";
import type { ProviderStore } from "../provider/store";
import { ToolExecutor } from "../tools/executor";
import {
  GuardedChatExecutor,
  listChatReadOnlyTools,
} from "./chatReadOnlyTools";
import {
  CHAT_READONLY_PROMPT_MARKER,
  composeChatReadonlySystemPrompt,
} from "./chatReadonlyPrompt";
import type {
  ChatControllerDeps,
  ChatControllerState,
  ChatPublicError,
  ChatSendResult,
  ChatTraceEvent,
  ChatTurnStatus,
} from "./types";

export class ChatController {
  private readonly store: ProviderStore;
  private readonly host: HostAdapter;
  private readonly fetchImpl?: ProviderFetch;
  private readonly maxRounds: number;
  private readonly composeSystemPrompt: (userMessage: string) => string;
  private readonly createProvider: ChatControllerDeps["createProvider"];
  private readonly onEvent?: (event: ChatTraceEvent) => void;

  private status: ChatControllerStatusState = "idle";
  private committed: AgentMessage[] = [];
  private lastTurnStatus?: ChatTurnStatus;
  private lastAssistantText?: string;
  private lastError?: ChatPublicError;
  private lastRunSummary?: ChatControllerState["lastRun"];
  private abortController: AbortController | null = null;

  constructor(deps: ChatControllerDeps) {
    this.store = deps.store;
    this.host = deps.host;
    this.fetchImpl = deps.fetchImpl;
    this.maxRounds = deps.maxRounds ?? 8;
    this.composeSystemPrompt =
      deps.composeSystemPrompt ??
      ((userMessage: string) =>
        composeChatReadonlySystemPrompt({
          routing: { content: userMessage },
        }));
    this.createProvider =
      deps.createProvider ??
      (() =>
        createStreamProviderFromStore(this.store, {
          fetchImpl: this.fetchImpl,
        }));
    this.onEvent = deps.onEvent;
  }

  getState(): ChatControllerState {
    return {
      status: this.status,
      messages: this.committed.slice(),
      lastTurnStatus: this.lastTurnStatus,
      lastAssistantText: this.lastAssistantText,
      lastError: this.lastError,
      lastRun: this.lastRunSummary,
    };
  }

  /** Clear committed history; only allowed while idle. */
  clear(): { ok: true } | { ok: false; error: string } {
    if (this.status !== "idle") {
      return { ok: false, error: "cannot clear while busy" };
    }
    this.committed = [];
    this.lastTurnStatus = undefined;
    this.lastAssistantText = undefined;
    this.lastError = undefined;
    this.lastRunSummary = undefined;
    return { ok: true };
  }

  stop(): void {
    if (this.status !== "running" || !this.abortController) return;
    this.status = "stopping";
    this.abortController.abort();
  }

  async send(userMessage: string): Promise<ChatSendResult> {
    const trimmed = typeof userMessage === "string" ? userMessage.trim() : "";
    if (!trimmed) {
      return this.preflightEnd("empty", { message: "empty message" });
    }
    if (this.status !== "idle") {
      return this.preflightEnd("busy", { message: "chat is busy" });
    }

    const providerResult = this.createProvider!();
    if (!providerResult.ok) {
      return this.preflightEnd("preflight_failed", {
        message: providerResult.error,
        kind: providerResult.kind,
      });
    }

    const systemPrompt = this.composeSystemPrompt(trimmed);
    const tools = listChatReadOnlyTools();
    const executor = new GuardedChatExecutor(new ToolExecutor(this.host));
    const ac = new AbortController();
    this.abortController = ac;
    this.status = "running";
    this.lastError = undefined;

    const history = this.committed.slice();
    const loop = new AgentLoop({
      provider: providerResult.provider,
      executor,
      systemPrompt,
      tools,
      maxRounds: this.maxRounds,
      signal: ac.signal,
      onEvent: (event) => this.projectEvent(event),
    });

    try {
      const runPromise = loop.run({ userMessage: trimmed, history });
      const result = await runPromise;
      // Commit always after a started run, regardless of terminal status.
      this.committed = result.messages.slice();
      this.lastAssistantText = result.assistantText;
      this.lastRunSummary = {
        status: result.status,
        rounds: result.rounds,
        assistantText: result.assistantText,
        lastFinishReason: result.lastFinishReason,
      };
      const turnStatus = mapRunStatus(result.status);
      this.lastTurnStatus = turnStatus;
      if (result.error) {
        this.lastError = {
          message: result.error.message,
          kind: result.error.kind,
          status: result.error.status,
          url: result.error.url,
        };
      }
      this.emit({ type: "turn_end", turnStatus });
      return {
        turnStatus,
        error: this.lastError,
        run: result,
      };
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      this.lastError = { message, kind: "provider" };
      this.lastTurnStatus = "failed";
      this.emit({ type: "turn_end", turnStatus: "failed" });
      return { turnStatus: "failed", error: this.lastError };
    } finally {
      this.status = "idle";
      this.abortController = null;
    }
  }

  /** Test/debug helper: expose whether prompt marker is expected. */
  static readonly readonlyMarker = CHAT_READONLY_PROMPT_MARKER;

  private preflightEnd(
    turnStatus: ChatTurnStatus,
    error: ChatPublicError,
  ): ChatSendResult {
    this.lastTurnStatus = turnStatus;
    this.lastError = error;
    this.emit({ type: "turn_end", turnStatus });
    return { turnStatus, error };
  }

  private projectEvent(event: LoopEvent): void {
    switch (event.type) {
      case "round_start":
        this.emit({ type: "round_start", round: event.round });
        break;
      case "text_delta":
        this.emit({
          type: "text_delta",
          delta: event.delta,
          round: event.round,
        });
        break;
      case "tool_call_parsed":
        this.emit({
          type: "tool_call_parsed",
          call: event.call,
          round: event.round,
        });
        break;
      case "tool_outcome":
        this.emit({
          type: "tool_outcome",
          toolCallId: event.toolCallId,
          outcome: event.outcome,
          round: event.round,
        });
        break;
      case "round_end":
        this.emit({
          type: "round_end",
          round: event.round,
          finishReason: event.finishReason,
          toolCallCount: event.toolCallCount,
        });
        break;
      case "run_end":
        this.emit({
          type: "run_end",
          status: event.result.status,
          rounds: event.result.rounds,
        });
        break;
      default:
        break;
    }
  }

  private emit(event: ChatTraceEvent): void {
    this.onEvent?.(event);
  }
}

type ChatControllerStatusState = ChatControllerState["status"];

function mapRunStatus(status: AgentRunResult["status"]): ChatTurnStatus {
  switch (status) {
    case "completed":
      return "completed";
    case "failed":
      return "failed";
    case "aborted":
      return "aborted";
    case "max_rounds":
      return "max_rounds";
    default:
      return "failed";
  }
}
