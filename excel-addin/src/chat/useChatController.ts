import { useCallback, useEffect, useRef, useState } from "react";
import {
  ChatController,
  type ApprovalRequest,
  type ChatControllerDeps,
  type ChatTraceEvent,
  type ChatTurnStatus,
} from "@shared/agentChat";
import type { AgentContentPart } from "@shared/agent";
import type { HostAdapter } from "@shared/host";
import type { ProviderStore } from "@shared/provider";
import {
  mapChatError,
  projectTraceEvent,
  type DisplayTraceItem,
  type DisplayTurn,
} from "./chatPresentation";

export interface ChatViewState {
  status: "idle" | "running" | "awaiting_approval" | "stopping";
  turns: DisplayTurn[];
  liveAssistant: string;
  bannerError?: string;
  pendingApproval: ApprovalRequest | null;
  canSend: boolean;
  canStop: boolean;
  canClear: boolean;
  canApprove: boolean;
  canReject: boolean;
}

export interface UseChatControllerOptions {
  store: ProviderStore;
  adapter: HostAdapter | null;
  createController?: (deps: ChatControllerDeps) => ChatController;
  fetchImpl?: ChatControllerDeps["fetchImpl"];
}

/** Result of send/retry so the composer can keep or clear the draft. */
export interface ChatSendOutcome {
  /** True when a turn was accepted into history (not preflight/busy/empty). */
  accepted: boolean;
  turnStatus?: ChatTurnStatus;
  /** User text that should stay in the composer when not accepted. */
  restoreText?: string;
  /** Final assistant text when a turn completed (for task result parsing). */
  assistantText?: string;
}

let turnSeq = 0;
function nextTurnId(): string {
  turnSeq += 1;
  return `turn-${turnSeq}`;
}

const PREFLIGHT_NO_TURN = new Set<ChatTurnStatus>([
  "preflight_failed",
  "busy",
  "empty",
]);

function appendTrace(
  turns: DisplayTurn[],
  turnId: string | null,
  item: DisplayTraceItem,
): DisplayTurn[] {
  if (!turnId) return turns;
  return turns.map((t) =>
    t.id === turnId ? { ...t, traces: [...t.traces, item] } : t,
  );
}

/**
 * Bridge ChatController → React view projection.
 * Per-instance generation token prevents disposed controllers from updating UI.
 */
export function useChatController(options: UseChatControllerOptions): {
  view: ChatViewState;
  send: (
    text: string,
    options?: { contentParts?: AgentContentPart[] },
  ) => Promise<ChatSendOutcome>;
  retry: (turnId: string) => Promise<ChatSendOutcome>;
  executeTool: (
    toolName: string,
    args: Record<string, unknown>,
    options?: { toolCallId?: string },
  ) => Promise<import("@shared/agentChat").ChatToolExecuteResult>;
  stop: () => void;
  clear: () => void;
  approve: (requestId?: string) => boolean;
  reject: (requestId?: string) => boolean;
  controller: ChatController | null;
} {
  const { store, adapter, createController, fetchImpl } = options;
  const [status, setStatus] = useState<ChatViewState["status"]>("idle");
  const [turns, setTurns] = useState<DisplayTurn[]>([]);
  const [liveAssistant, setLiveAssistant] = useState("");
  const [bannerError, setBannerError] = useState<string | undefined>();
  const [pendingApproval, setPendingApproval] =
    useState<ApprovalRequest | null>(null);
  const [controller, setController] = useState<ChatController | null>(null);
  // Sync for retry without stale closure.
  // (assigned each render intentionally)

  const eventSeq = useRef(0);
  const activeTurnId = useRef<string | null>(null);
  const turnsRef = useRef<DisplayTurn[]>([]);
  turnsRef.current = turns;
  const generationRef = useRef(0);
  const disposedRef = useRef(false);
  const controllerRef = useRef<ChatController | null>(null);
  const createControllerRef = useRef(createController);
  const fetchImplRef = useRef(fetchImpl);
  createControllerRef.current = createController;
  fetchImplRef.current = fetchImpl;

  const isLive = useCallback((gen: number) => {
    return !disposedRef.current && generationRef.current === gen;
  }, []);

  const handleEvent = useCallback(
    (gen: number, event: ChatTraceEvent) => {
      if (!isLive(gen)) return;

      if (event.type === "approval_needed") {
        setPendingApproval({ ...event.request });
        setStatus((prev) =>
          prev === "stopping" ? "stopping" : "awaiting_approval",
        );
        eventSeq.current += 1;
        const item = projectTraceEvent(event, eventSeq.current);
        if (item) {
          setTurns((prev) => appendTrace(prev, activeTurnId.current, item));
        }
        return;
      }

      if (event.type === "approval_resolved") {
        setPendingApproval((prev) =>
          prev && prev.requestId === event.requestId ? null : prev,
        );
        setStatus((prev) => {
          if (prev === "stopping") return "stopping";
          if (
            event.decision === "approved" ||
            event.decision === "rejected" ||
            event.decision === "cancelled"
          ) {
            return "running";
          }
          return prev;
        });
        eventSeq.current += 1;
        const item = projectTraceEvent(event, eventSeq.current);
        if (item) {
          setTurns((prev) => appendTrace(prev, activeTurnId.current, item));
        }
        return;
      }

      if (event.type === "text_delta") {
        setLiveAssistant((prev) => prev + event.delta);
        const id = activeTurnId.current;
        if (id) {
          setTurns((prev) =>
            prev.map((t) =>
              t.id === id
                ? {
                    ...t,
                    assistantText: t.assistantText + event.delta,
                    pending: true,
                  }
                : t,
            ),
          );
        }
        return;
      }
      if (event.type === "turn_end") return;
      eventSeq.current += 1;
      const item = projectTraceEvent(event, eventSeq.current);
      if (!item) return;
      setTurns((prev) => appendTrace(prev, activeTurnId.current, item));
    },
    [isLive],
  );

  useEffect(() => {
    disposedRef.current = false;
    const gen = ++generationRef.current;
    activeTurnId.current = null;

    if (!adapter) {
      controllerRef.current = null;
      setController(null);
      setPendingApproval(null);
      return () => {
        disposedRef.current = true;
        generationRef.current += 1;
      };
    }

    const factory =
      createControllerRef.current ??
      ((deps: ChatControllerDeps) => new ChatController(deps));
    const c = factory({
      store,
      host: adapter,
      fetchImpl: fetchImplRef.current,
      onEvent: (e) => handleEvent(gen, e),
    });
    controllerRef.current = c;
    setController(c);
    setStatus("idle");
    setLiveAssistant("");
    setPendingApproval(null);
    activeTurnId.current = null;

    return () => {
      disposedRef.current = true;
      if (generationRef.current === gen) {
        generationRef.current += 1;
      }
      try {
        c.stop();
      } catch {
        /* ignore */
      }
      if (controllerRef.current === c) {
        controllerRef.current = null;
      }
      setController((prev) => (prev === c ? null : prev));
      setPendingApproval(null);
      activeTurnId.current = null;
    };
  }, [adapter, store, handleEvent]);

  const send = useCallback(
    async (
      text: string,
      options?: { contentParts?: AgentContentPart[] },
    ): Promise<ChatSendOutcome> => {
      const c = controllerRef.current;
      if (!c) {
        return { accepted: false, restoreText: text };
      }
      if (c.getState().status !== "idle") {
        return { accepted: false, turnStatus: "busy", restoreText: text };
      }
      const gen = generationRef.current;
      const trimmed = typeof text === "string" ? text.trim() : "";
      if (!trimmed) {
        if (isLive(gen)) {
          setBannerError(mapChatError(undefined, "empty"));
        }
        return { accepted: false, turnStatus: "empty", restoreText: text };
      }

      const id = nextTurnId();
      activeTurnId.current = id;
      if (isLive(gen)) {
        setLiveAssistant("");
        setBannerError(undefined);
        setPendingApproval(null);
        setTurns((prev) => [
          ...prev,
          {
            id,
            userText: trimmed,
            assistantText: "",
            pending: true,
            traces: [],
          },
        ]);
        setStatus("running");
      }

      const result = await c.send(trimmed, options);
      if (!isLive(gen)) {
        return { accepted: false, turnStatus: result.turnStatus };
      }

      const errText = mapChatError(result.error, result.turnStatus);
      setBannerError(errText);
      setPendingApproval(null);

      if (PREFLIGHT_NO_TURN.has(result.turnStatus) || result.run == null) {
        setTurns((prev) => prev.filter((t) => t.id !== id));
        setLiveAssistant("");
        setStatus("idle");
        if (activeTurnId.current === id) activeTurnId.current = null;
        return {
          accepted: false,
          turnStatus: result.turnStatus,
          restoreText: trimmed,
        };
      }

      setTurns((prev) =>
        prev.map((t) =>
          t.id === id
            ? {
                ...t,
                pending: false,
                turnStatus: result.turnStatus,
                assistantText:
                  result.run?.assistantText ?? t.assistantText ?? "",
                errorText: errText,
              }
            : t,
        ),
      );
      setLiveAssistant("");
      setStatus("idle");
      if (activeTurnId.current === id) activeTurnId.current = null;
      return {
        accepted: true,
        turnStatus: result.turnStatus,
        assistantText:
          result.run?.assistantText ??
          controllerRef.current?.getState().lastAssistantText,
      };
    },
    [isLive],
  );

  const retry = useCallback(
    async (turnId: string): Promise<ChatSendOutcome> => {
      const c = controllerRef.current;
      if (!c || c.getState().status !== "idle") {
        return { accepted: false, turnStatus: "busy" };
      }
      const snapshot = turnsRef.current.find((t) => t.id === turnId);
      const text = (snapshot?.userText ?? "").trim();
      if (!text) {
        return { accepted: false, turnStatus: "empty" };
      }
      // New send appends a new turn; does not mutate controller history until accepted.
      return send(text);
    },
    [send],
  );

  const stop = useCallback(() => {
    const c = controllerRef.current;
    if (!c) return;
    c.stop();
    if (!disposedRef.current) setStatus("stopping");
  }, []);

  const clear = useCallback(() => {
    const c = controllerRef.current;
    if (!c) return;
    const r = c.clear();
    if (!r.ok) {
      if (!disposedRef.current) setBannerError(r.error);
      return;
    }
    if (disposedRef.current) return;
    setTurns([]);
    setLiveAssistant("");
    setBannerError(undefined);
    setPendingApproval(null);
    setStatus("idle");
    activeTurnId.current = null;
  }, []);

  const approve = useCallback((requestId?: string) => {
    const c = controllerRef.current;
    if (!c) return false;
    return c.approve(requestId);
  }, []);


  const executeTool = useCallback(
    async (
      toolName: string,
      args: Record<string, unknown>,
      options?: { toolCallId?: string },
    ) => {
      const c = controllerRef.current;
      if (!c) {
        return {
          ok: false,
          tool: toolName,
          error: "controller not ready",
        };
      }
      if (c.getState().status !== "idle") {
        return { ok: false, tool: toolName, error: "chat is busy" };
      }
      const gen = generationRef.current;
      if (!disposedRef.current) {
        setBannerError(undefined);
        setPendingApproval(null);
        setStatus("running");
      }
      try {
        const result = await c.executeTool(toolName, args, options);
        if (!isLive(gen)) {
          return result;
        }
        setPendingApproval(null);
        setStatus("idle");
        if (!result.ok && result.error) {
          setBannerError(result.error);
        }
        return result;
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        if (isLive(gen)) {
          setPendingApproval(null);
          setStatus("idle");
          setBannerError(message);
        }
        return { ok: false, tool: toolName, error: message };
      }
    },
    [isLive],
  );

  const reject = useCallback((requestId?: string) => {
    const c = controllerRef.current;
    if (!c) return false;
    return c.reject(requestId);
  }, []);

  const busy =
    status === "running" ||
    status === "awaiting_approval" ||
    status === "stopping";
  const canDecide =
    !!controller &&
    status === "awaiting_approval" &&
    pendingApproval != null;

  return {
    controller,
    view: {
      status,
      turns,
      liveAssistant,
      bannerError,
      pendingApproval,
      canSend: !!controller && !busy,
      canStop: status === "running" || status === "awaiting_approval",
      canClear: !!controller && status === "idle",
      canApprove: canDecide,
      canReject: canDecide,
    },
    send,
    retry,
    executeTool,
    stop,
    clear,
    approve,
    reject,
  };
}

export type { ChatTurnStatus, DisplayTraceItem, ApprovalRequest };
