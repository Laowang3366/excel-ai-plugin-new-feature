import type { ChatState } from "./chatStore";

type TurnStartBaseState = Pick<ChatState, "activeThreadId" | "stoppedThreadIds">;

type TurnStartPatch = Pick<
  ChatState,
  | "isStreaming"
  | "streamingContent"
  | "streamingReasoning"
  | "activeStreamingRound"
  | "turnStatus"
  | "activeClientId"
  | "stoppedThreadIds"
  | "error"
> &
  Partial<Pick<ChatState, "compactionNotice" | "lastInterruptContext">>;

function removeActiveStoppedThread(state: TurnStartBaseState): ChatState["stoppedThreadIds"] {
  if (!state.activeThreadId) {
    return state.stoppedThreadIds;
  }
  return Object.fromEntries(
    Object.entries(state.stoppedThreadIds).filter(([id]) => id !== state.activeThreadId),
  );
}

export function buildTurnStartPatch(
  state: TurnStartBaseState,
  clientId: string,
  extraPatch: Partial<Pick<ChatState, "compactionNotice" | "lastInterruptContext">> = {},
): TurnStartPatch {
  return {
    isStreaming: true,
    streamingContent: "",
    streamingReasoning: "",
    activeStreamingRound: null,
    turnStatus: "in_progress",
    activeClientId: clientId,
    stoppedThreadIds: removeActiveStoppedThread(state),
    error: null,
    ...extraPatch,
  };
}
