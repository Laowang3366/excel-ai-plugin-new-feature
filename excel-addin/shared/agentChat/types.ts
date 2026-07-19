import type {
  AgentMessage,
  AgentRunResult,
  AgentToolOutcome,
  LoopEvent,
  ParsedToolCall,
} from "../agent/types";
import type { ProviderFetch } from "../provider/client";
import type { ProviderStore } from "../provider/store";
import type { HostAdapter } from "../host/hostAdapter";
import type { CreateStreamProviderResult } from "../provider/createStreamProvider";
import type { ApprovalDecision, ApprovalRequest } from "./approvalGate";

export type ChatControllerStatus =
  | "idle"
  | "running"
  | "awaiting_approval"
  | "stopping";

export type ChatTurnStatus =
  | "completed"
  | "failed"
  | "aborted"
  | "max_rounds"
  | "preflight_failed"
  | "busy"
  | "empty";

export interface ChatPublicError {
  message: string;
  kind?: string;
  status?: number;
  url?: string;
}

export interface ChatControllerState {
  status: ChatControllerStatus;
  /** Committed history used as the next turn's history source of truth. */
  messages: AgentMessage[];
  lastTurnStatus?: ChatTurnStatus;
  lastAssistantText?: string;
  lastError?: ChatPublicError;
  lastRun?: Pick<
    AgentRunResult,
    "status" | "rounds" | "assistantText" | "lastFinishReason"
  >;
  /** Public-only pending approval (no raw args). */
  pendingApproval?: ApprovalRequest | null;
}

/** Projected loop + approval events (no stream-level tool arg deltas). */
export type ChatTraceEvent =
  | { type: "round_start"; round: number }
  | { type: "text_delta"; delta: string; round: number }
  | { type: "tool_call_parsed"; call: ParsedToolCall; round: number }
  | {
      type: "tool_outcome";
      toolCallId: string;
      outcome: AgentToolOutcome;
      round: number;
    }
  | {
      type: "round_end";
      round: number;
      finishReason: string;
      toolCallCount: number;
    }
  | { type: "run_end"; status: AgentRunResult["status"]; rounds: number }
  | { type: "turn_end"; turnStatus: ChatTurnStatus }
  | { type: "approval_needed"; request: ApprovalRequest }
  | {
      type: "approval_resolved";
      requestId: string;
      decision: ApprovalDecision;
      request: ApprovalRequest;
    };

export interface ChatControllerDeps {
  store: ProviderStore;
  host: HostAdapter;
  fetchImpl?: ProviderFetch;
  maxRounds?: number;
  /** Test seam: override system prompt composition. */
  composeSystemPrompt?: (userMessage: string) => string;
  /** Test seam: override provider factory (defaults to createStreamProviderFromStore). */
  createProvider?: () => CreateStreamProviderResult;
  onEvent?: (event: ChatTraceEvent) => void;
}

export interface ChatSendResult {
  turnStatus: ChatTurnStatus;
  error?: ChatPublicError;
  run?: AgentRunResult;
}

export type { LoopEvent, ApprovalRequest, ApprovalDecision };
