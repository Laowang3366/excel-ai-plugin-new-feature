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
  status: "idle" | "running" | "stopping";
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

/**
 * Bridge ChatController → React view projection.
 * onEvent is held in a ref so StrictMode remounts do not recreate controller mid-turn.
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
  const onEventRef = useRef<(event: ChatTraceEvent) => void>(() => {});
  const createControllerRef = useRef(createController);
  const fetchImplRef = useRef(fetchImpl);
  createControllerRef.current = createController;
  fetchImplRef.current = fetchImpl;

  onEventRef.current = (event: ChatTraceEvent) => {
    if (event.type === "text_delta") {
      setLiveAssistant((prev) => prev + event.delta);
      const id = activeTurnId.current;
      if (id) {
        setTurns((prev) =>
          prev.map((t) =>
            t.id === id
              ? { ...t, assistantText: t.assistantText + event.delta, pending: true }
              : t,
          ),
        );
      }
      return;
    }
    if (event.type === "turn_end") {
      return;
    }
    eventSeq.current += 1;
    const item = projectTraceEvent(event, eventSeq.current);
    if (!item) return;
    const id = activeTurnId.current;
    if (!id) return;
    setTurns((prev) =>
      prev.map((t) => (t.id === id ? { ...t, traces: [...t.traces, item] } : t)),
    );
  };

  useEffect(() => {
    if (!adapter) {
      setController(null);
      return;
    }
    const factory =
      createControllerRef.current ??
      ((deps: ChatControllerDeps) => new ChatController(deps));
    const c = factory({
      store,
      host: adapter,
      fetchImpl: fetchImplRef.current,
      onEvent: (e) => onEventRef.current(e),
    });
    setController(c);
    return () => {
      // No dispose API; drop reference.
      setController(null);
    };
  }, [adapter, store]);

  const send = useCallback(
    async (text: string) => {
      if (!controller) return;
      if (controller.getState().status !== "idle") return;
      const trimmed = text.trim();
      if (!trimmed) {
        setBannerError(mapChatError(undefined, "empty"));
        return;
      }
      const id = nextTurnId();
      activeTurnId.current = id;
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
      const result = await controller.send(trimmed);
      const errText = mapChatError(result.error, result.turnStatus);
      setBannerError(errText);
      setTurns((prev) =>
        prev.map((t) =>
          t.id === id
            ? {
                ...t,
                pending: false,
                turnStatus: result.turnStatus,
                assistantText:
                  result.run?.assistantText ??
                  t.assistantText ??
                  "",
                errorText: errText,
              }
            : t,
        ),
      );
      setLiveAssistant("");
      setStatus("idle");
      activeTurnId.current = null;
    },
    [controller],
  );

  const stop = useCallback(() => {
    if (!controller) return;
    controller.stop();
    setStatus("stopping");
  }, [controller]);

  const clear = useCallback(() => {
    if (!controller) return;
    const r = controller.clear();
    if (!r.ok) {
      setBannerError(r.error);
      return;
    }
    setTurns([]);
    setLiveAssistant("");
    setBannerError(undefined);
    setStatus("idle");
  }, [controller]);

  const busy = status === "running" || status === "stopping";
  return {
    controller,
    view: {
      status,
      turns,
      liveAssistant,
      bannerError,
      canSend: !!controller && !busy,
      canStop: status === "running",
      canClear: !!controller && status === "idle",
    },
    send,
    stop,
    clear,
  };
}

export type { ChatTurnStatus, DisplayTraceItem };
