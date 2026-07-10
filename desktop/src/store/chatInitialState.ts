import type { ChatState } from "./chatStore";

export function createInitialChatState(): ChatState {
  return {
    messages: [],
    isStreaming: false,
    streamingContent: "",
    streamingReasoning: "",
    activeStreamingRound: null,
    showReasoning: true,
    reasoningExpanded: {},
    activeTurnId: null,
    activeThreadId: null,
    activeClientId: null,
    runningThreadIds: {},
    pendingInterruptThreadIds: {},
    stoppedThreadIds: {},
    turnStatus: "idle",
    lastInterruptContext: null,
    tokenUsage: null,
    contextUsage: null,
    compactionNotice: null,
    error: null,
    threads: [],
    pendingToolCall: null,
    pendingComposerFiles: [],
    pendingFolderId: null,
  };
}

export function createClearedMessagesPatch(): Partial<ChatState> {
  return {
    messages: [],
    streamingContent: "",
    streamingReasoning: "",
    activeStreamingRound: null,
    activeTurnId: null,
    turnStatus: "idle",
    lastInterruptContext: null,
    tokenUsage: null,
    contextUsage: null,
    compactionNotice: null,
    error: null,
  };
}
