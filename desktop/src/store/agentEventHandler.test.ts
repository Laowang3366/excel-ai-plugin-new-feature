import { describe, expect, it } from "vitest";
import type { ChatState } from "./chatStore";
import type { AgentEvent, TurnItem } from "../electronApi";
import { handleAgentEvent } from "./agentEventHandler";

function makeState(patch: Partial<ChatState> = {}): ChatState {
  return {
    messages: [],
    isStreaming: true,
    streamingContent: "",
    streamingReasoning: "",
    activeStreamingRound: 0,
    showReasoning: true,
    reasoningExpanded: {},
    activeTurnId: "turn-1",
    activeThreadId: "thread-1",
    activeClientId: null,
    runningThreadIds: {},
    pendingInterruptThreadIds: {},
    stoppedThreadIds: {},
    turnStatus: "in_progress",
    lastInterruptContext: null,
    tokenUsage: null,
    contextUsage: null,
    compactionNotice: null,
    error: null,
    threads: [],
    pendingToolCall: null,
    pendingComposerFiles: [],
    pendingFolderId: null,
    ...patch,
  };
}

function applyPatches(state: ChatState, patches: Array<Partial<ChatState>>): ChatState {
  return Object.assign({}, state, ...patches);
}

describe("handleAgentEvent", () => {
  it("marks the UI as streaming when a queued turn starts from the backend", () => {
    const current = makeState({
      isStreaming: false,
      turnStatus: "completed",
      error: "previous error",
    });

    const next = applyPatches(
      current,
      handleAgentEvent(
        { type: "turn_started", turnId: "turn-queued", threadId: "thread-1" } as AgentEvent,
        current,
        []
      )
    );

    expect(next).toMatchObject({
      activeTurnId: "turn-queued",
      turnStatus: "in_progress",
      isStreaming: true,
      error: null,
    });
  });

  it("ignores agent events from a different thread", () => {
    const current = makeState({
      activeThreadId: "thread-2",
      messages: [],
      isStreaming: false,
      turnStatus: "idle",
    });
    const foreignItem: TurnItem = {
      type: "assistant_message",
      id: "msg-foreign",
      content: "旧会话仍在输出",
      phase: "final",
      timestamp: 2000,
    };

    const patches = handleAgentEvent(
      { type: "item_completed", threadId: "thread-1", item: foreignItem } as AgentEvent,
      current,
      []
    );

    expect(patches).toEqual([]);
  });

  it("tracks a different running thread without projecting its messages into the active view", () => {
    const current = makeState({
      activeThreadId: "thread-2",
      messages: [],
      isStreaming: false,
      turnStatus: "idle",
    });

    const patches = handleAgentEvent(
      { type: "turn_started", turnId: "turn-foreign", threadId: "thread-1" } as AgentEvent,
      current,
      []
    );

    expect(patches).toEqual([{
      runningThreadIds: { "thread-1": true },
      stoppedThreadIds: {},
    }]);
  });

  it("ignores old thread items while the active view is a blank new conversation", () => {
    const current = makeState({
      activeThreadId: null,
      activeClientId: null,
      messages: [],
      isStreaming: false,
      turnStatus: "idle",
    });
    const oldItem: TurnItem = {
      type: "tool_call",
      id: "tool-old",
      toolName: "python_execute",
      arguments: {},
      status: "completed",
      timestamp: 2000,
    };

    const patches = handleAgentEvent(
      { type: "item_completed", threadId: "thread-old", item: oldItem } as AgentEvent,
      current,
      []
    );

    expect(patches).toEqual([]);
  });

  it("binds a pending new conversation when the matching client turn starts", () => {
    const current = makeState({
      activeThreadId: null,
      activeClientId: "client-new",
      messages: [],
      isStreaming: true,
      turnStatus: "in_progress",
    });

    const patches = handleAgentEvent(
      {
        type: "turn_started",
        turnId: "turn-new",
        threadId: "thread-new",
        clientId: "client-new",
      } as AgentEvent,
      current,
      []
    );

    const next = applyPatches(current, patches);
    expect(next).toMatchObject({
      activeThreadId: "thread-new",
      activeTurnId: "turn-new",
      turnStatus: "in_progress",
      isStreaming: true,
    });
  });

  it("does not bind a blank new conversation to another client's running thread", () => {
    const current = makeState({
      activeThreadId: null,
      activeClientId: "client-new",
      messages: [],
      isStreaming: true,
      turnStatus: "in_progress",
    });

    const patches = handleAgentEvent(
      {
        type: "turn_started",
        turnId: "turn-old",
        threadId: "thread-old",
        clientId: "client-old",
      } as AgentEvent,
      current,
      []
    );

    expect(patches).toEqual([{
      runningThreadIds: { "thread-old": true },
      stoppedThreadIds: {},
    }]);
  });

  it("clears a different running thread when it completes", () => {
    const current = makeState({
      activeThreadId: "thread-2",
      runningThreadIds: { "thread-1": true, "thread-2": true },
      isStreaming: false,
      turnStatus: "idle",
    });

    const patches = handleAgentEvent(
      { type: "turn_completed", turnId: "turn-foreign", threadId: "thread-1" } as AgentEvent,
      current,
      []
    );

    expect(patches).toEqual([{
      runningThreadIds: { "thread-2": true },
      pendingInterruptThreadIds: {},
    }]);
  });

  it("clears active turn id when the active thread is interrupted", () => {
    const current = makeState({
      activeThreadId: "thread-1",
      activeTurnId: "turn-running",
      isStreaming: true,
      turnStatus: "in_progress",
    });

    const next = applyPatches(
      current,
      handleAgentEvent(
        { type: "turn_interrupted", turnId: "turn-running", threadId: "thread-1" } as AgentEvent,
        current,
        []
      )
    );

    expect(next).toMatchObject({
      activeTurnId: null,
      isStreaming: false,
      turnStatus: "interrupted",
    });
  });

  it("freezes streaming reasoning before a tool call so the next reasoning segment starts clean", () => {
    const toolCall: TurnItem = {
      type: "tool_call",
      id: "tool-1",
      toolName: "selection.get",
      arguments: {},
      status: "pending",
      timestamp: 2000,
    };
    const current = makeState({
      streamingReasoning: "第一段思考",
      streamingContent: "",
    });

    const next = applyPatches(
      current,
      handleAgentEvent({ type: "item_started", item: toolCall } as AgentEvent, current, [])
    );

    expect(next.streamingReasoning).toBe("");
    expect(next.messages).toHaveLength(2);
    expect(next.messages[0]).toMatchObject({
      type: "reasoning",
      rawContent: ["第一段思考"],
    });
    expect(next.messages[1]).toBe(toolCall);
  });

  it("replaces a frozen reasoning segment when the completed reasoning item arrives", () => {
    const frozen: TurnItem = {
      type: "reasoning",
      id: "streaming-reasoning-turn-1-0-tool-1",
      summaryText: [],
      rawContent: ["第一段思考"],
      timestamp: 1000,
    };
    const toolCall: TurnItem = {
      type: "tool_call",
      id: "tool-1",
      toolName: "selection.get",
      arguments: {},
      status: "completed",
      timestamp: 2000,
    };
    const completed: TurnItem = {
      type: "reasoning",
      id: "reasoning-1",
      summaryText: [],
      rawContent: ["第一段", "思考"],
      timestamp: 3000,
    };
    const current = makeState({
      messages: [frozen, toolCall],
      streamingReasoning: "第一段思考",
    });

    const next = applyPatches(
      current,
      handleAgentEvent({ type: "item_completed", item: completed } as AgentEvent, current, [])
    );

    expect(next.streamingReasoning).toBe("");
    expect(next.messages).toHaveLength(2);
    expect(next.messages[0]).toBe(completed);
    expect(next.messages[1]).toBe(toolCall);
  });

  it("clears stale streaming buffers when the final assistant message is completed", () => {
    const finalMessage: TurnItem = {
      type: "assistant_message",
      id: "msg-final",
      content: "处理完成",
      phase: "final",
      timestamp: 4000,
    };
    const current = makeState({
      streamingContent: "",
      streamingReasoning: "迟到的旧思考片段",
      activeStreamingRound: 3,
    });

    const next = applyPatches(
      current,
      handleAgentEvent({ type: "item_completed", item: finalMessage } as AgentEvent, current, [])
    );

    expect(next.streamingContent).toBe("");
    expect(next.streamingReasoning).toBe("");
    expect(next.activeStreamingRound).toBeNull();
    expect(next.messages).toEqual([finalMessage]);
  });

  it("does not clear newer round streaming buffers for an older commentary completion", () => {
    const commentary: TurnItem = {
      type: "assistant_message",
      id: "msg-commentary",
      content: "上一轮工具前说明",
      phase: "commentary",
      timestamp: 3000,
    };
    const current = makeState({
      streamingContent: "下一轮正在输出",
      streamingReasoning: "下一轮正在思考",
      activeStreamingRound: 2,
    });

    const next = applyPatches(
      current,
      handleAgentEvent({ type: "item_completed", item: commentary } as AgentEvent, current, [])
    );

    expect(next.streamingContent).toBe("下一轮正在输出");
    expect(next.streamingReasoning).toBe("下一轮正在思考");
    expect(next.activeStreamingRound).toBe(2);
    expect(next.messages).toEqual([commentary]);
  });

  it("does not clear newer round reasoning for an older completed reasoning item", () => {
    const completed: TurnItem = {
      type: "reasoning",
      id: "reasoning-old",
      summaryText: [],
      rawContent: ["上一轮思考"],
      timestamp: 3000,
    };
    const current = makeState({
      streamingReasoning: "下一轮思考",
      activeStreamingRound: 2,
    });

    const next = applyPatches(
      current,
      handleAgentEvent({ type: "item_completed", item: completed } as AgentEvent, current, [])
    );

    expect(next.streamingReasoning).toBe("下一轮思考");
    expect(next.messages).toEqual([completed]);
  });

  it("shows compaction progress items in the message list", () => {
    const progress: TurnItem = {
      type: "compact_progress",
      id: "compact-1",
      reason: "auto_token_limit",
      status: "running",
      message: "正在压缩上下文...",
      timestamp: 1000,
    };
    const current = makeState({ messages: [] });

    const next = applyPatches(
      current,
      handleAgentEvent({ type: "item_started", item: progress } as AgentEvent, current, [])
    );

    expect(next.messages).toEqual([progress]);
  });

  it("stores a visible notice when compaction starts", () => {
    const current = makeState();

    const next = applyPatches(
      current,
      handleAgentEvent({
        type: "thread_compact_started",
        threadId: "thread-1",
        params: {
          reason: "auto_token_limit",
          itemCount: 8,
          tokensBefore: 120000,
          tokenThreshold: 100000,
          retryCount: 2,
          timestamp: 1000,
        },
      } as AgentEvent, current, [])
    );

    expect(next.compactionNotice).toContain("失败最多重试 2 次");
  });
});
