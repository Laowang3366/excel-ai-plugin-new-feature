import { beforeEach, describe, expect, it, vi } from "vitest";

// @MOCK_INTERFACE: mirrors ipcApi.agent/thread methods used by chatStore actions and stream listeners.
const ipcMocks = vi.hoisted(() => ({
  startTurn: vi.fn(),
  continueTurn: vi.fn(),
  enqueueTurn: vi.fn(),
  interrupt: vi.fn(),
  resumeThread: vi.fn(),
  listThreads: vi.fn(),
  loadThread: vi.fn(),
  deleteThread: vi.fn(),
  newThread: vi.fn(),
  onEvent: vi.fn(() => () => {}),
  onStreamDelta: vi.fn(() => () => {}),
}));

vi.mock("../services/ipcApi", () => ({
  ipcApi: {
    agent: {
      startTurn: ipcMocks.startTurn,
      continueTurn: ipcMocks.continueTurn,
      enqueueTurn: ipcMocks.enqueueTurn,
      interrupt: ipcMocks.interrupt,
      onEvent: ipcMocks.onEvent,
      onStreamDelta: ipcMocks.onStreamDelta,
    },
    thread: {
      resume: ipcMocks.resumeThread,
      list: ipcMocks.listThreads,
      load: ipcMocks.loadThread,
      delete: ipcMocks.deleteThread,
      newThread: ipcMocks.newThread,
    },
  },
}));

import { mergeBufferedStreamDeltas, useChatStore } from "./chatStore";

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (reason?: unknown) => void;
  const promise = new Promise<T>((resolvePromise, rejectPromise) => {
    resolve = resolvePromise;
    reject = rejectPromise;
  });
  return { promise, resolve, reject };
}

describe("chatStore sendMessage", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    ipcMocks.enqueueTurn.mockResolvedValue({ success: true, queued: true, queueSize: 1 });
    ipcMocks.listThreads.mockResolvedValue([]);
    useChatStore.setState({
      isStreaming: false,
      turnStatus: "idle",
      activeThreadId: "thread-1",
      error: null,
    });
  });

  it("queues new user input while the agent is already streaming", async () => {
    useChatStore.setState({
      isStreaming: true,
      turnStatus: "in_progress",
      activeThreadId: "thread-1",
      error: null,
    });

    await useChatStore.getState().sendMessage("等一下，用 Sheet2 的数据");

    expect(ipcMocks.enqueueTurn).toHaveBeenCalledWith(expect.objectContaining({
      content: "等一下，用 Sheet2 的数据",
      attachments: undefined,
      threadId: "thread-1",
      isResume: false,
    }));
    expect(ipcMocks.enqueueTurn.mock.calls[0][0].clientId).toMatch(/^client-/);
    expect(ipcMocks.startTurn).not.toHaveBeenCalled();
    expect(useChatStore.getState().error).toBeNull();
    expect(useChatStore.getState().isStreaming).toBe(true);
  });

  it("does not enqueue into an unbound new conversation before thread id arrives", async () => {
    useChatStore.setState({
      isStreaming: true,
      turnStatus: "in_progress",
      activeThreadId: null,
      activeClientId: "client-pending",
      error: null,
    });

    await useChatStore.getState().sendMessage("second message too early");

    expect(ipcMocks.enqueueTurn).not.toHaveBeenCalled();
    expect(ipcMocks.startTurn).not.toHaveBeenCalled();
    expect(useChatStore.getState().error).toBe("会话正在创建中，请等待连接完成后再发送");
    expect(useChatStore.getState().isStreaming).toBe(true);
  });

  it("returns the thread id created for a new conversation", async () => {
    const startTurn = deferred<{ success: boolean; threadId: string }>();
    ipcMocks.startTurn.mockReturnValue(startTurn.promise);
    useChatStore.setState({
      isStreaming: false,
      turnStatus: "idle",
      activeThreadId: null,
      activeClientId: null,
      error: null,
    });

    const send = useChatStore.getState().sendMessage("创建新会话");
    await Promise.resolve();
    useChatStore.setState({ activeThreadId: "thread-existing", activeClientId: null });
    startTurn.resolve({ success: true, threadId: "thread-created" });

    await expect(send).resolves.toBe("thread-created");
    expect(useChatStore.getState().activeThreadId).toBe("thread-existing");
  });

  it("returns null when resuming an existing thread throws", async () => {
    ipcMocks.resumeThread.mockRejectedValue(new Error("resume unavailable"));

    await expect(useChatStore.getState().sendMessage("继续处理")).resolves.toBeNull();

    expect(ipcMocks.startTurn).not.toHaveBeenCalled();
    expect(useChatStore.getState().error).toBe("resume unavailable");
  });
});

describe("chatStore switchThread", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    useChatStore.setState({
      activeThreadId: null,
      activeTurnId: null,
      activeClientId: null,
      messages: [],
      isStreaming: false,
      runningThreadIds: {},
      pendingInterruptThreadIds: {},
      stoppedThreadIds: {},
      turnStatus: "idle",
      error: null,
    });
  });

  it("keeps the newest thread when an older load finishes last", async () => {
    const loadA = deferred<{ items: Array<Record<string, unknown>> }>();
    const loadB = deferred<{ items: Array<Record<string, unknown>> }>();
    ipcMocks.loadThread.mockImplementation((threadId: string) =>
      threadId === "thread-a" ? loadA.promise : loadB.promise
    );

    const switchA = useChatStore.getState().switchThread("thread-a");
    const switchB = useChatStore.getState().switchThread("thread-b");

    loadB.resolve({
      items: [{
        type: "assistant_message",
        id: "message-b",
        content: "会话 B",
        timestamp: 2,
      }],
    });
    await switchB;

    loadA.resolve({
      items: [{
        type: "assistant_message",
        id: "message-a",
        content: "会话 A",
        timestamp: 1,
      }],
    });
    await switchA;

    expect(useChatStore.getState().activeThreadId).toBe("thread-b");
    expect(useChatStore.getState().messages).toMatchObject([
      { id: "message-b", content: "会话 B" },
    ]);
  });

  it("ignores an older load error after a newer switch succeeds", async () => {
    const loadA = deferred<{ items: Array<Record<string, unknown>> }>();
    const loadB = deferred<{ items: Array<Record<string, unknown>> }>();
    ipcMocks.loadThread.mockImplementation((threadId: string) =>
      threadId === "thread-a" ? loadA.promise : loadB.promise
    );

    const switchA = useChatStore.getState().switchThread("thread-a");
    const switchB = useChatStore.getState().switchThread("thread-b");

    loadB.resolve({ items: [] });
    await switchB;
    loadA.reject(new Error("A 加载失败"));
    await switchA;

    expect(useChatStore.getState().activeThreadId).toBe("thread-b");
    expect(useChatStore.getState().error).toBeNull();
  });

  it("ignores an older load after creating a new thread", async () => {
    const loadA = deferred<{ items: Array<Record<string, unknown>> }>();
    ipcMocks.loadThread.mockReturnValue(loadA.promise);
    ipcMocks.newThread.mockResolvedValue({ success: true });
    useChatStore.setState({
      activeThreadId: "thread-current",
      messages: [{
        type: "assistant_message",
        id: "message-current",
        content: "当前会话",
        timestamp: 0,
      }],
    });

    const switchA = useChatStore.getState().switchThread("thread-a");
    await useChatStore.getState().createNewThread("D:\\work");

    loadA.resolve({
      items: [{
        type: "assistant_message",
        id: "message-a",
        content: "会话 A",
        timestamp: 1,
      }],
    });
    await switchA;

    expect(useChatStore.getState().activeThreadId).toBeNull();
    expect(useChatStore.getState().messages).toEqual([]);
    expect(useChatStore.getState().pendingFolderId).toBe("D:\\work");
  });

  it("ignores an older load after deleting the current thread", async () => {
    const loadA = deferred<{ items: Array<Record<string, unknown>> }>();
    ipcMocks.loadThread.mockReturnValue(loadA.promise);
    ipcMocks.deleteThread.mockResolvedValue({ success: true });
    ipcMocks.listThreads.mockResolvedValue([]);
    useChatStore.setState({
      activeThreadId: "thread-current",
      messages: [{
        type: "assistant_message",
        id: "message-current",
        content: "当前会话",
        timestamp: 0,
      }],
    });

    const switchA = useChatStore.getState().switchThread("thread-a");
    await useChatStore.getState().deleteThread("thread-current");

    loadA.resolve({
      items: [{
        type: "assistant_message",
        id: "message-a",
        content: "会话 A",
        timestamp: 1,
      }],
    });
    await switchA;

    expect(useChatStore.getState().activeThreadId).toBeNull();
    expect(useChatStore.getState().messages).toEqual([]);
  });

  it("ignores a load after deleting its pending target thread", async () => {
    const loadA = deferred<{ items: Array<Record<string, unknown>> }>();
    ipcMocks.loadThread.mockReturnValue(loadA.promise);
    ipcMocks.deleteThread.mockResolvedValue({ success: true });
    ipcMocks.listThreads.mockResolvedValue([]);
    useChatStore.setState({
      activeThreadId: "thread-current",
      messages: [{
        type: "assistant_message",
        id: "message-current",
        content: "当前会话",
        timestamp: 0,
      }],
    });

    const switchA = useChatStore.getState().switchThread("thread-a");
    await useChatStore.getState().deleteThread("thread-a");

    loadA.resolve({
      items: [{
        type: "assistant_message",
        id: "message-a",
        content: "会话 A",
        timestamp: 1,
      }],
    });
    await switchA;

    expect(useChatStore.getState().activeThreadId).toBe("thread-current");
    expect(useChatStore.getState().messages).toMatchObject([
      { id: "message-current", content: "当前会话" },
    ]);
  });

  it("allows a pending switch to finish after deleting an unrelated thread", async () => {
    const loadA = deferred<{ items: Array<Record<string, unknown>> }>();
    ipcMocks.loadThread.mockReturnValue(loadA.promise);
    ipcMocks.deleteThread.mockResolvedValue({ success: true });
    ipcMocks.listThreads.mockResolvedValue([]);
    useChatStore.setState({
      activeThreadId: "thread-current",
      messages: [{
        type: "assistant_message",
        id: "message-current",
        content: "当前会话",
        timestamp: 0,
      }],
    });

    const switchA = useChatStore.getState().switchThread("thread-a");
    await useChatStore.getState().deleteThread("thread-b");

    loadA.resolve({
      items: [{
        type: "assistant_message",
        id: "message-a",
        content: "会话 A",
        timestamp: 1,
      }],
    });
    await switchA;

    expect(useChatStore.getState().activeThreadId).toBe("thread-a");
    expect(useChatStore.getState().messages).toMatchObject([
      { id: "message-a", content: "会话 A" },
    ]);
  });

  it("keeps a newer switch when an older new-thread request finishes last", async () => {
    const createRequest = deferred<{ success: boolean }>();
    ipcMocks.newThread.mockReturnValue(createRequest.promise);
    ipcMocks.loadThread.mockResolvedValue({
      items: [{
        type: "assistant_message",
        id: "message-b",
        content: "会话 B",
        timestamp: 2,
      }],
    });

    const createThread = useChatStore.getState().createNewThread("D:\\work");
    const switchB = useChatStore.getState().switchThread("thread-b");
    await switchB;

    createRequest.resolve({ success: true });
    await createThread;

    expect(useChatStore.getState().activeThreadId).toBe("thread-b");
    expect(useChatStore.getState().messages).toMatchObject([
      { id: "message-b", content: "会话 B" },
    ]);
  });

  it("keeps a newer switch when an older delete finishes last", async () => {
    const deleteRequest = deferred<{ success: boolean }>();
    ipcMocks.deleteThread.mockReturnValue(deleteRequest.promise);
    ipcMocks.listThreads.mockResolvedValue([]);
    ipcMocks.loadThread.mockResolvedValue({
      items: [{
        type: "assistant_message",
        id: "message-b",
        content: "会话 B",
        timestamp: 2,
      }],
    });
    useChatStore.setState({
      activeThreadId: "thread-current",
      messages: [{
        type: "assistant_message",
        id: "message-current",
        content: "当前会话",
        timestamp: 0,
      }],
    });

    const deleteCurrent = useChatStore.getState().deleteThread("thread-current");
    const switchB = useChatStore.getState().switchThread("thread-b");
    await switchB;

    deleteRequest.resolve({ success: true });
    await deleteCurrent;

    expect(useChatStore.getState().activeThreadId).toBe("thread-b");
    expect(useChatStore.getState().messages).toMatchObject([
      { id: "message-b", content: "会话 B" },
    ]);
  });

  it("clears a thread selected after its deletion request began", async () => {
    const deleteRequest = deferred<{ success: boolean }>();
    ipcMocks.deleteThread.mockReturnValue(deleteRequest.promise);
    ipcMocks.listThreads.mockResolvedValue([]);
    ipcMocks.loadThread.mockResolvedValue({
      items: [{
        type: "assistant_message",
        id: "message-b",
        content: "会话 B",
        timestamp: 2,
      }],
    });
    useChatStore.setState({
      activeThreadId: "thread-current",
      messages: [{
        type: "assistant_message",
        id: "message-current",
        content: "当前会话",
        timestamp: 0,
      }],
    });

    const deleteB = useChatStore.getState().deleteThread("thread-b");
    await useChatStore.getState().switchThread("thread-b");

    deleteRequest.resolve({ success: true });
    await deleteB;

    expect(useChatStore.getState().activeThreadId).toBeNull();
    expect(useChatStore.getState().messages).toEqual([]);
  });

  it("cancels a pending switch when its target is deleted", async () => {
    const deleteRequest = deferred<{ success: boolean }>();
    const loadB = deferred<{ items: Array<Record<string, unknown>> }>();
    ipcMocks.deleteThread.mockReturnValue(deleteRequest.promise);
    ipcMocks.listThreads.mockResolvedValue([]);
    ipcMocks.loadThread.mockReturnValue(loadB.promise);
    useChatStore.setState({
      activeThreadId: "thread-current",
      messages: [{
        type: "assistant_message",
        id: "message-current",
        content: "当前会话",
        timestamp: 0,
      }],
    });

    const deleteB = useChatStore.getState().deleteThread("thread-b");
    const switchB = useChatStore.getState().switchThread("thread-b");

    deleteRequest.resolve({ success: true });
    await deleteB;
    loadB.resolve({
      items: [{
        type: "assistant_message",
        id: "message-b",
        content: "会话 B",
        timestamp: 2,
      }],
    });
    await switchB;

    expect(useChatStore.getState().activeThreadId).toBe("thread-current");
    expect(useChatStore.getState().messages).toMatchObject([
      { id: "message-current", content: "当前会话" },
    ]);
  });
});

describe("chatStore stream delta handling", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    useChatStore.setState({
      messages: [],
      isStreaming: false,
      streamingContent: "",
      streamingReasoning: "",
      activeStreamingRound: null,
      turnStatus: "completed",
      activeThreadId: "thread-1",
      activeClientId: null,
      error: null,
    });
  });

  it("ignores late stream deltas after the turn has completed", () => {
    useChatStore.getState().handleStreamDelta({
      delta: "迟到的思考",
      itemType: "reasoning",
      roundId: 2,
      threadId: "thread-1",
    });

    expect(useChatStore.getState().streamingReasoning).toBe("");
    expect(useChatStore.getState().streamingContent).toBe("");
    expect(useChatStore.getState().activeStreamingRound).toBeNull();
  });

  it("merges adjacent buffered stream deltas before updating store state", () => {
    expect(mergeBufferedStreamDeltas([
      { delta: "思", itemType: "reasoning", roundId: 1, threadId: "thread-1" },
      { delta: "考", itemType: "reasoning", roundId: 1, threadId: "thread-1" },
      { delta: "正文", itemType: "assistant_message", roundId: 1, threadId: "thread-1" },
      { delta: "新轮", itemType: "reasoning", roundId: 2, threadId: "thread-1" },
    ])).toEqual([
      { delta: "思考", itemType: "reasoning", roundId: 1, threadId: "thread-1" },
      { delta: "正文", itemType: "assistant_message", roundId: 1, threadId: "thread-1" },
      { delta: "新轮", itemType: "reasoning", roundId: 2, threadId: "thread-1" },
    ]);
  });
});

describe("chatStore interruptTurn", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    ipcMocks.listThreads.mockResolvedValue([]);
    useChatStore.setState({
      isStreaming: true,
      turnStatus: "in_progress",
      activeThreadId: "thread-1",
      activeTurnId: "turn-1",
      activeClientId: null,
      runningThreadIds: { "thread-1": true },
      pendingInterruptThreadIds: {},
      stoppedThreadIds: {},
      error: null,
    });
  });

  it("keeps the conversation locked until the interrupt IPC confirms completion", async () => {
    let resolveInterrupt!: (value: { success: boolean }) => void;
    ipcMocks.interrupt.mockReturnValue(new Promise((resolve) => {
      resolveInterrupt = resolve;
    }));

    const interruptResult = useChatStore.getState().interruptTurn();
    useChatStore.getState().handleAgentEvent({
      type: "turn_interrupted",
      turnId: "turn-1",
      threadId: "thread-1",
    });

    expect(useChatStore.getState().runningThreadIds).toEqual({ "thread-1": true });
    expect(useChatStore.getState().stoppedThreadIds).toEqual({});
    expect(useChatStore.getState().isStreaming).toBe(true);
    expect(useChatStore.getState().turnStatus).toBe("in_progress");

    resolveInterrupt({ success: true });
    await interruptResult;

    expect(useChatStore.getState().runningThreadIds).toEqual({});
    expect(useChatStore.getState().stoppedThreadIds).toEqual({ "thread-1": true });
    expect(useChatStore.getState().isStreaming).toBe(false);
    expect(useChatStore.getState().turnStatus).toBe("interrupted");
  });

  it("keeps the running state when interrupt IPC fails", async () => {
    ipcMocks.interrupt.mockResolvedValue({ success: false });

    await useChatStore.getState().interruptTurn();

    expect(useChatStore.getState().runningThreadIds).toEqual({ "thread-1": true });
    expect(useChatStore.getState().stoppedThreadIds).toEqual({});
    expect(useChatStore.getState().isStreaming).toBe(true);
    expect(useChatStore.getState().turnStatus).toBe("in_progress");
    expect(useChatStore.getState().error).toBe("停止当前任务失败，请稍后重试");
  });

  it("unlocks stale running state when the backend reports no active agent", async () => {
    ipcMocks.interrupt.mockResolvedValue({
      success: false,
      error: "没有正在运行的 Agent",
    });

    await useChatStore.getState().interruptTurn();

    expect(useChatStore.getState().runningThreadIds).toEqual({});
    expect(useChatStore.getState().stoppedThreadIds).toEqual({ "thread-1": true });
    expect(useChatStore.getState().isStreaming).toBe(false);
    expect(useChatStore.getState().activeTurnId).toBeNull();
    expect(useChatStore.getState().turnStatus).toBe("interrupted");
    expect(useChatStore.getState().error).toBeNull();
  });

  it("keeps the running state when interrupt IPC throws", async () => {
    ipcMocks.interrupt.mockRejectedValue(new Error("ipc unavailable"));

    await useChatStore.getState().interruptTurn();

    expect(useChatStore.getState().runningThreadIds).toEqual({ "thread-1": true });
    expect(useChatStore.getState().stoppedThreadIds).toEqual({});
    expect(useChatStore.getState().isStreaming).toBe(true);
    expect(useChatStore.getState().turnStatus).toBe("in_progress");
    expect(useChatStore.getState().error).toBe("停止当前任务失败，请稍后重试");
  });

  it("does not unlock a different thread selected while interrupt is pending", async () => {
    let resolveInterrupt!: (value: { success: boolean }) => void;
    ipcMocks.interrupt.mockReturnValue(new Promise((resolve) => {
      resolveInterrupt = resolve;
    }));

    const interruptResult = useChatStore.getState().interruptTurn();
    useChatStore.setState({
      activeThreadId: "thread-2",
      activeTurnId: "turn-2",
      isStreaming: true,
      streamingContent: "thread-2 output",
      turnStatus: "in_progress",
      runningThreadIds: { "thread-1": true, "thread-2": true },
    });

    resolveInterrupt({ success: true });
    await interruptResult;

    expect(useChatStore.getState().activeThreadId).toBe("thread-2");
    expect(useChatStore.getState().activeTurnId).toBe("turn-2");
    expect(useChatStore.getState().isStreaming).toBe(true);
    expect(useChatStore.getState().streamingContent).toBe("thread-2 output");
    expect(useChatStore.getState().turnStatus).toBe("in_progress");
    expect(useChatStore.getState().runningThreadIds).toEqual({ "thread-2": true });
    expect(useChatStore.getState().stoppedThreadIds).toEqual({ "thread-1": true });
  });
});

describe("chatStore loadThreads", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    useChatStore.setState({
      threads: [],
      runningThreadIds: {},
      stoppedThreadIds: {},
      activeThreadId: null,
      isStreaming: false,
      turnStatus: "idle",
      error: null,
    });
  });

  it("reconciles running thread ids from thread metadata", async () => {
    useChatStore.setState({
      runningThreadIds: {
        "thread-stale": true,
      },
    });
    ipcMocks.listThreads.mockResolvedValue([
      {
        threadId: "thread-running",
        preview: "运行中",
        modelProvider: "test",
        createdAt: 1,
        updatedAt: 3,
        activeTurnId: "turn-running",
      },
      {
        threadId: "thread-stale",
        preview: "已完成",
        modelProvider: "test",
        createdAt: 1,
        updatedAt: 2,
        lastTurnStatus: "completed",
      },
    ]);

    await useChatStore.getState().loadThreads();

    expect(useChatStore.getState().runningThreadIds).toEqual({
      "thread-running": true,
    });
  });

  it("does not revive a user-stopped thread from stale in-progress metadata", async () => {
    useChatStore.setState({
      stoppedThreadIds: { "thread-stopped": true },
      runningThreadIds: {},
    });
    ipcMocks.listThreads.mockResolvedValue([
      {
        threadId: "thread-stopped",
        preview: "旧运行态",
        modelProvider: "test",
        createdAt: 1,
        updatedAt: 2,
        activeTurnId: "turn-stale",
        lastTurnStatus: "in_progress",
      },
    ]);

    await useChatStore.getState().loadThreads();

    expect(useChatStore.getState().runningThreadIds).toEqual({});
  });
});

describe("chatStore createNewThread", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    ipcMocks.listThreads.mockResolvedValue([]);
    ipcMocks.newThread.mockResolvedValue({ success: true });
    useChatStore.setState({
      messages: [{ type: "user_message", id: "msg-1", content: "正在处理", timestamp: 1 } as any],
      isStreaming: true,
      turnStatus: "in_progress",
      activeThreadId: "thread-running",
      runningThreadIds: { "thread-running": true },
      error: null,
    });
  });

  it("does not clear the active conversation while a turn is streaming", async () => {
    await useChatStore.getState().createNewThread();

    expect(ipcMocks.newThread).not.toHaveBeenCalled();
    expect(useChatStore.getState().messages).toHaveLength(1);
    expect(useChatStore.getState().isStreaming).toBe(true);
    expect(useChatStore.getState().turnStatus).toBe("in_progress");
    expect(useChatStore.getState().activeThreadId).toBe("thread-running");
    expect(useChatStore.getState().error).toBe("当前会话正在执行，请等待完成或停止后再新建会话");
  });

  it("blocks new thread creation while another thread is running in the background", async () => {
    useChatStore.setState({
      isStreaming: false,
      turnStatus: "idle",
      activeThreadId: "thread-idle",
      runningThreadIds: { "thread-background": true },
      error: null,
    });

    await useChatStore.getState().createNewThread();

    expect(ipcMocks.newThread).not.toHaveBeenCalled();
    expect(useChatStore.getState().activeThreadId).toBe("thread-idle");
    expect(useChatStore.getState().error).toBe("当前会话正在执行，请等待完成或停止后再新建会话");
  });
});
