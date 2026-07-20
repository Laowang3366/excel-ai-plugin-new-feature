import { AgentLoop } from "../agent/agentLoop";
import type { AgentMessage, AgentRunResult, LoopEvent } from "../agent/types";
import type { HostAdapter } from "../host/hostAdapter";
import { createStreamProviderFromStore } from "../provider/createStreamProvider";
import type { ProviderFetch } from "../provider/client";
import type { ProviderStore } from "../provider/store";
import { ToolExecutor } from "../tools/executor";
import { ApprovalGate, type ApprovalRequest } from "./approvalGate";
import { ApprovingToolExecutor } from "./approvingToolExecutor";
import {
  DEFAULT_PERMISSION_MODE,
  type PermissionMode,
} from "./approvalPolicy";
import { getBrowserPermissionModeStore } from "./permissionModeStore";
import {
  CHAT_APPROVAL_PROMPT_MARKER,
  composeChatApprovalSystemPrompt,
} from "./chatApprovalPrompt";
import { listChatTools } from "./chatToolPolicy";
import { resolveChatPromptRuntimeContext } from "./promptRuntimeContext";
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
  private readonly composeSystemPrompt: NonNullable<
    ChatControllerDeps["composeSystemPrompt"]
  >;
  private readonly createProvider: ChatControllerDeps["createProvider"];
  private readonly getPermissionMode: () => PermissionMode;
  private readonly onEvent?: (event: ChatTraceEvent) => void;

  private status: ChatControllerState["status"] = "idle";
  private committed: AgentMessage[] = [];
  private lastTurnStatus?: ChatTurnStatus;
  private lastAssistantText?: string;
  private lastError?: ChatPublicError;
  private lastRunSummary?: ChatControllerState["lastRun"];
  private abortController: AbortController | null = null;
  private gate: ApprovalGate | null = null;
  private pendingApproval: ApprovalRequest | null = null;
  private currentToolCallId: string | undefined;
  private currentRound: number | undefined;
  private abortListener: (() => void) | null = null;
  private unsubGate: (() => void) | null = null;

  constructor(deps: ChatControllerDeps) {
    this.store = deps.store;
    this.host = deps.host;
    this.fetchImpl = deps.fetchImpl;
    this.maxRounds = deps.maxRounds ?? 8;
    // Assign before compose default so prompt + ApprovingToolExecutor share one getter.
    this.getPermissionMode =
      deps.getPermissionMode ??
      (() => {
        try {
          return getBrowserPermissionModeStore().get();
        } catch {
          return DEFAULT_PERMISSION_MODE;
        }
      });
    this.composeSystemPrompt =
      deps.composeSystemPrompt ??
      ((userMessage, runtimeContext) =>
        composeChatApprovalSystemPrompt({
          routing: { content: userMessage },
          ...runtimeContext,
          // Re-read each turn so mode switches apply on the next send (same source as executor).
          permissionMode: this.getPermissionMode(),
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
      pendingApproval: this.pendingApproval
        ? { ...this.pendingApproval }
        : null,
    };
  }

  clear(): { ok: true } | { ok: false; error: string } {
    if (this.status !== "idle") {
      return { ok: false, error: "cannot clear while busy" };
    }
    this.committed = [];
    this.lastTurnStatus = undefined;
    this.lastAssistantText = undefined;
    this.lastError = undefined;
    this.lastRunSummary = undefined;
    this.pendingApproval = null;
    return { ok: true };
  }

  stop(): void {
    if (
      (this.status !== "running" && this.status !== "awaiting_approval") ||
      !this.abortController
    ) {
      return;
    }
    this.status = "stopping";
    this.gate?.cancelAll("chat stop");
    this.abortController.abort();
  }

  approve(requestId?: string): boolean {
    const gate = this.gate;
    const pending = this.pendingApproval;
    if (!gate || !pending) return false;
    const id = requestId ?? pending.requestId;
    return gate.approve(id);
  }

  reject(requestId?: string): boolean {
    const gate = this.gate;
    const pending = this.pendingApproval;
    if (!gate || !pending) return false;
    const id = requestId ?? pending.requestId;
    return gate.reject(id);
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

    // Create AbortController before any await so stop() works during host probe.
    this.status = "running";
    const ac = new AbortController();
    this.abortController = ac;

    let systemPrompt: string;
    try {
      const runtimeContext = await resolveChatPromptRuntimeContext(this.host);
      if (ac.signal.aborted) {
        return this.finishAbortedBeforeLoop();
      }
      systemPrompt = this.composeSystemPrompt(trimmed, runtimeContext);
    } catch (error) {
      if (ac.signal.aborted) {
        return this.finishAbortedBeforeLoop();
      }
      this.abortController = null;
      this.status = "idle";
      return this.preflightEnd("preflight_failed", {
        message: error instanceof Error ? error.message : String(error),
      });
    }
    if (ac.signal.aborted) {
      return this.finishAbortedBeforeLoop();
    }

    const tools = listChatTools();
    const gate = new ApprovalGate();
    this.gate = gate;
    this.pendingApproval = null;
    this.currentToolCallId = undefined;
    this.currentRound = undefined;

    this.abortListener = () => {
      gate.cancelAll("aborted");
    };
    ac.signal.addEventListener("abort", this.abortListener);

    this.unsubGate = gate.subscribe((event) => {
      if (event.type === "requested") {
        this.pendingApproval = { ...event.request };
        if (this.status !== "stopping") {
          this.status = "awaiting_approval";
        }
        this.emit({ type: "approval_needed", request: { ...event.request } });
        return;
      }
      // resolved
      if (this.pendingApproval?.requestId === event.requestId) {
        this.pendingApproval = null;
      }
      if (
        (event.decision === "approved" || event.decision === "rejected") &&
        this.status === "awaiting_approval"
      ) {
        this.status = "running";
      }
      this.emit({
        type: "approval_resolved",
        requestId: event.requestId,
        decision: event.decision,
        request: { ...event.request },
      });
    });

    const executor = new ApprovingToolExecutor(
      new ToolExecutor(this.host),
      gate,
      () => ({
        toolCallId: this.currentToolCallId,
        round: this.currentRound,
      }),
      () => this.getPermissionMode(),
    );

    this.lastError = undefined;

    const history = this.committed.slice();
    const activeProvider = this.store.getActive();
    const loop = new AgentLoop({
      provider: providerResult.provider,
      executor,
      systemPrompt,
      tools,
      maxRounds: this.maxRounds,
      signal: ac.signal,
      onEvent: (event) => this.projectEvent(event),
      contextWindowSize: activeProvider?.contextWindowSize,
    });

    try {
      const result = await loop.run({ userMessage: trimmed, history });
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
      this.gate?.cancelAll("turn end");
      if (this.abortListener && this.abortController) {
        this.abortController.signal.removeEventListener(
          "abort",
          this.abortListener,
        );
      }
      this.unsubGate?.();
      this.unsubGate = null;
      this.abortListener = null;
      this.gate = null;
      this.abortController = null;
      this.pendingApproval = null;
      this.currentToolCallId = undefined;
      this.currentRound = undefined;
      this.status = "idle";
    }
  }

  static readonly approvalMarker = CHAT_APPROVAL_PROMPT_MARKER;

  /** Stop during host probe / prompt compose: no provider call, idle again. */
  private finishAbortedBeforeLoop(): ChatSendResult {
    this.abortController = null;
    this.status = "idle";
    return this.preflightEnd("aborted", { message: "aborted", kind: "aborted" });
  }

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
        this.currentRound = event.round;
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
        this.currentToolCallId = event.call.id;
        this.currentRound = event.round;
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
        this.currentToolCallId = undefined;
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
