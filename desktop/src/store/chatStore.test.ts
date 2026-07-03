import { beforeEach, describe, expect, it, vi } from "vitest";

const ipcMocks = vi.hoisted(() => ({
  startTurn: vi.fn(),
  continueTurn: vi.fn(),
  enqueueTurn: vi.fn(),
  interrupt: vi.fn(),
  resumeThread: vi.fn(),
  listThreads: vi.fn(),
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
    },
  },
}));

import { useChatStore } from "./chatStore";

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
