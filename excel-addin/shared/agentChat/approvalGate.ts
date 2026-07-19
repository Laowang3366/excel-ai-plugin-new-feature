import type { RiskLevel } from "../tools/types";
import type { ArgsPreview } from "./approvalPreview";

export type ApprovalDecision = "approved" | "rejected" | "cancelled";

export interface ApprovalRequest {
  requestId: string;
  toolCallId?: string;
  round?: number;
  name: string;
  riskLevel: RiskLevel;
  destructive: boolean;
  argsPreview: ArgsPreview;
  impactHint: string;
  createdAt: number;
}

export type ApprovalGateEvent =
  | { type: "requested"; request: ApprovalRequest }
  | {
      type: "resolved";
      requestId: string;
      decision: ApprovalDecision;
      request: ApprovalRequest;
    };

type Pending = {
  request: ApprovalRequest;
  resolve: (decision: "approved" | "rejected") => void;
  reject: (error: Error) => void;
};

function abortError(message: string): Error {
  if (typeof DOMException !== "undefined") {
    return new DOMException(message, "AbortError");
  }
  return Object.assign(new Error(message), { name: "AbortError" });
}

let reqSeq = 0;
function nextId(): string {
  reqSeq += 1;
  return `apr-${Date.now().toString(36)}-${reqSeq}`;
}

/**
 * Per-turn approval gate. Public state never holds raw tool arguments.
 * cancelAll settles pending as AbortError and arms the gate for this turn.
 */
export class ApprovalGate {
  private pending: Pending | null = null;
  private cancelled = false;
  private cancelReason = "approval cancelled";
  private readonly listeners = new Set<(event: ApprovalGateEvent) => void>();

  subscribe(listener: (event: ApprovalGateEvent) => void): () => void {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  }

  getPending(): ApprovalRequest | null {
    return this.pending ? { ...this.pending.request } : null;
  }

  isCancelled(): boolean {
    return this.cancelled;
  }

  request(
    input: Omit<ApprovalRequest, "requestId" | "createdAt"> & {
      requestId?: string;
      createdAt?: number;
    },
  ): Promise<"approved" | "rejected"> {
    if (this.cancelled) {
      return Promise.reject(abortError(this.cancelReason));
    }
    if (this.pending) {
      return Promise.reject(
        new Error("approval gate already has a pending request"),
      );
    }
    const request: ApprovalRequest = {
      requestId: input.requestId ?? nextId(),
      toolCallId: input.toolCallId,
      round: input.round,
      name: input.name,
      riskLevel: input.riskLevel,
      destructive: input.destructive,
      argsPreview: input.argsPreview,
      impactHint: input.impactHint,
      createdAt: input.createdAt ?? Date.now(),
    };
    return new Promise<"approved" | "rejected">((resolve, reject) => {
      this.pending = { request, resolve, reject };
      this.emit({ type: "requested", request: { ...request } });
    });
  }

  approve(requestId: string): boolean {
    const p = this.pending;
    if (!p || p.request.requestId !== requestId) return false;
    this.pending = null;
    p.resolve("approved");
    this.emit({
      type: "resolved",
      requestId,
      decision: "approved",
      request: { ...p.request },
    });
    return true;
  }

  reject(requestId: string): boolean {
    const p = this.pending;
    if (!p || p.request.requestId !== requestId) return false;
    this.pending = null;
    p.resolve("rejected");
    this.emit({
      type: "resolved",
      requestId,
      decision: "rejected",
      request: { ...p.request },
    });
    return true;
  }

  cancelAll(reason = "approval cancelled"): void {
    this.cancelled = true;
    this.cancelReason = reason;
    const p = this.pending;
    this.pending = null;
    if (!p) return;
    p.reject(abortError(reason));
    this.emit({
      type: "resolved",
      requestId: p.request.requestId,
      decision: "cancelled",
      request: { ...p.request },
    });
  }

  private emit(event: ApprovalGateEvent): void {
    for (const listener of this.listeners) {
      try {
        listener(event);
      } catch {
        /* ignore subscriber errors */
      }
    }
  }
}
