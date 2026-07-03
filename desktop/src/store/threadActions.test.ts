import { beforeEach, describe, expect, test, vi } from "vitest";
import type { ChatState } from "./chatStore";

const mocks = vi.hoisted(() => ({
  interrupt: vi.fn(),
  resume: vi.fn(),
  load: vi.fn(),
  newThread: vi.fn(),
}));

vi.mock("../services/ipcApi", () => ({
  ipcApi: {
    agent: {
      interrupt: mocks.interrupt,
    },
    thread: {
      resume: mocks.resume,
      load: mocks.load,
      newThread: mocks.newThread,
    },
  },
}));

import { createNewThread, switchThread } from "./threadActions";

function neverSettles() {
  return new Promise<{ success: boolean }>(() => {});
}

describe("threadActions", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  test("switches thread without interrupting a running turn", async () => {
    mocks.interrupt.mockImplementation(neverSettles);
    mocks.resume.mockResolvedValue({ success: true });
    mocks.load.mockResolvedValue({ items: [] });

    const result = await switchThread("thread-2", { isStreaming: true } as ChatState);

    expect(mocks.interrupt).not.toHaveBeenCalled();
    expect(mocks.resume).not.toHaveBeenCalled();
    expect(mocks.load).toHaveBeenCalledWith("thread-2");
    expect(result.patches[0]).toMatchObject({
      activeThreadId: "thread-2",
      messages: [],
      isStreaming: false,
      turnStatus: "idle",
    });
  });

  test("restores streaming controls when switching back to a running thread", async () => {
    mocks.load.mockResolvedValue({ items: [] });

    const result = await switchThread("thread-1", {
      isStreaming: false,
      runningThreadIds: { "thread-1": true },
      stoppedThreadIds: {},
      threads: [{
        threadId: "thread-1",
        preview: "正在执行的会话",
        modelProvider: "test",
        createdAt: 1,
        updatedAt: 2,
        activeTurnId: "turn-running",
      }],
    } as unknown as ChatState);

    expect(result.patches[0]).toMatchObject({
      activeThreadId: "thread-1",
      isStreaming: true,
      activeTurnId: "turn-running",
      turnStatus: "in_progress",
    });
  });

  test("uses freshly loaded metadata instead of stale running state when switching thread", async () => {
    mocks.load.mockResolvedValue({
      metadata: {
        threadId: "thread-1",
        preview: "已中断的会话",
        modelProvider: "test",
        createdAt: 1,
        updatedAt: 3,
        lastTurnStatus: "interrupted",
      },
      turns: [],
    });

    const result = await switchThread("thread-1", {
      isStreaming: false,
      runningThreadIds: { "thread-1": true },
      stoppedThreadIds: { "thread-1": true },
      threads: [{
        threadId: "thread-1",
        preview: "旧状态",
        modelProvider: "test",
        createdAt: 1,
        updatedAt: 2,
        activeTurnId: "turn-stale",
        lastTurnStatus: "in_progress",
      }],
    } as unknown as ChatState);

    expect(result.patches[0]).toMatchObject({
      activeThreadId: "thread-1",
      isStreaming: false,
      activeTurnId: null,
      turnStatus: "idle",
    });
  });

  test("keeps a live running marker even when loaded rollout metadata looks interrupted", async () => {
    mocks.load.mockResolvedValue({
      metadata: {
        threadId: "thread-1",
        preview: "后台运行中",
        modelProvider: "test",
        createdAt: 1,
        updatedAt: 3,
        lastTurnStatus: "interrupted",
      },
      turns: [],
    });

    const result = await switchThread("thread-1", {
      isStreaming: false,
      runningThreadIds: { "thread-1": true },
      stoppedThreadIds: {},
      threads: [{
        threadId: "thread-1",
        preview: "运行中",
        modelProvider: "test",
        createdAt: 1,
        updatedAt: 2,
      }],
    } as unknown as ChatState);

    expect(result.patches[0]).toMatchObject({
      activeThreadId: "thread-1",
      isStreaming: true,
      turnStatus: "in_progress",
    });
  });

  test("blocks creating a new thread while a conversation is running", async () => {
    mocks.interrupt.mockImplementation(neverSettles);
    mocks.newThread.mockResolvedValue({ success: true });

    const result = await createNewThread(undefined, true);

    expect(mocks.interrupt).not.toHaveBeenCalled();
    expect(mocks.newThread).not.toHaveBeenCalled();
    expect(result.patches).toEqual([]);
    expect(result.error).toBe("当前会话正在执行，请等待完成或停止后再新建会话");
  });

  test("creates a new empty thread when idle", async () => {
    mocks.newThread.mockResolvedValue({ success: true });

    const result = await createNewThread("D:\\work", false);

    expect(mocks.newThread).toHaveBeenCalledWith("D:\\work");
    expect(result.error).toBeUndefined();
    expect(result.patches[0]).toMatchObject({
      messages: [],
      isStreaming: false,
      activeStreamingRound: null,
      activeThreadId: null,
      turnStatus: "idle",
      pendingFolderId: "D:\\work",
    });
  });
});
