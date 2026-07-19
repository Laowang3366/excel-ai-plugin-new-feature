import { useCallback, useEffect, useRef, useState } from "react";
import {
  ChatController,
  type ChatControllerDeps,
  type ChatTraceEvent,
  type ChatTurnStatus,
} from "@shared/agentChat";
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
  canSend: boolean;
  canStop: boolean;
  canClear: boolean;
}

export interface UseChatControllerOptions {
  store: ProviderStore;
  adapter: HostAdapter | null;
  /** Test seam; production default uses ChatController. */
  createController?: (deps: ChatControllerDeps) => ChatController;
  fetchImpl?: ChatControllerDeps["fetchImpl"];
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

/**
 * Bridge ChatController → React view projection.
 * Per-instance generation token prevents disposed controllers from updating UI.
 */
export function useChatController(options: UseChatControllerOptions): {
  view: ChatViewState;
  send: (text: string) => Promise<void>;
  stop: () => void;
  clear: () => void;
  controller: ChatController | null;
} {
  const { store, adapter, createController, fetchImpl } = options;
  const [status, setStatus] = useState<ChatViewState["status"]>("idle");
  const [turns, setTurns] = useState<DisplayTurn[]>([]);
  const [liveAssistant, setLiveAssistant] = useState("");
  const [bannerError, setBannerError] = useState<string | undefined>();
  const [controller, setController] = useState<ChatController | null>(null);

  const eventSeq = useRef(0);
  const activeTurnId = useRef<string | null>(null);
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
        setStatus("awaiting_approval");
        return;
      }
      if (event.type === "approval_resolved") {
        setStatus((prev) =>
          prev === "stopping" ? "stopping" : "running",
        );
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
      const id = activeTurnId.current;
      if (!id) return;
      setTurns((prev) =>
        prev.map((t) =>
          t.id === id ? { ...t, traces: [...t.traces, item] } : t,
        ),
      );
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
    // Fresh controller instance always starts idle for UI projection.
    setStatus("idle");
    setLiveAssistant("");
    activeTurnId.current = null;

    return () => {
      // 1) Mark disposed so old onEvent becomes no-op immediately.
      disposedRef.current = true;
      // 2) Invalidate generation so late events from this instance are dropped.
      if (generationRef.current === gen) {
        generationRef.current += 1;
      }
      // 3) Stop this instance only.
      try {
        c.stop();
      } catch {
        /* ignore */
      }
      // 4) Clear refs for this instance without clobbering a newer controller.
      if (controllerRef.current === c) {
        controllerRef.current = null;
      }
      setController((prev) => (prev === c ? null : prev));
      activeTurnId.current = null;
    };
  }, [adapter, store, handleEvent]);

  const send = useCallback(async (text: string) => {
    const c = controllerRef.current;
    if (!c) return;
    if (c.getState().status !== "idle") return;
    const gen = generationRef.current;
    const trimmed = typeof text === "string" ? text.trim() : "";
    if (!trimmed) {
      if (!isLive(gen)) return;
      setBannerError(mapChatError(undefined, "empty"));
      return;
    }

    const id = nextTurnId();
    // Optimistic turn — removed if preflight fails before run.
    activeTurnId.current = id;
    if (isLive(gen)) {
      setLiveAssistant("");
      setBannerError(undefined);
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

    const result = await c.send(trimmed);
    if (!isLive(gen)) return;

    const errText = mapChatError(result.error, result.turnStatus);
    setBannerError(errText);

    if (PREFLIGHT_NO_TURN.has(result.turnStatus) || result.run == null) {
      // Preflight / busy / empty: drop optimistic bubble; banner only.
      setTurns((prev) => prev.filter((t) => t.id !== id));
      setLiveAssistant("");
      setStatus("idle");
      if (activeTurnId.current === id) activeTurnId.current = null;
      return;
    }

    setTurns((prev) =>
      prev.map((t) =>
        t.id === id
          ? {
              ...t,
              pending: false,
              turnStatus: result.turnStatus,
              assistantText: result.run?.assistantText ?? t.assistantText ?? "",
              errorText: errText,
            }
          : t,
      ),
    );
    setLiveAssistant("");
    setStatus("idle");
    if (activeTurnId.current === id) activeTurnId.current = null;
  }, [isLive]);

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
    setStatus("idle");
    activeTurnId.current = null;
  }, []);

  const busy =
    status === "running" ||
    status === "awaiting_approval" ||
    status === "stopping";
  return {
    controller,
    view: {
      status,
      turns,
      liveAssistant,
      bannerError,
      canSend: !!controller && !busy,
      canStop: status === "running" || status === "awaiting_approval",
      canClear: !!controller && status === "idle",
    },
    send,
    stop,
    clear,
  };
}

export type { ChatTurnStatus, DisplayTraceItem };
