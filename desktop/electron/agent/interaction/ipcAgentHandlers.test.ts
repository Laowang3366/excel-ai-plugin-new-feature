import { describe, expect, it, vi } from "vitest";

import {
  enqueueTurnForIpc,
  listThreadsForIpc,
  prepareAgentForStartTurn,
  prepareNewThreadForIpc,
} from "./ipcAgentHandlers";

describe("listThreadsForIpc", () => {
  it("uses sqlite thread snapshots before falling back to JSONL session scanning", async () => {
    const sessionStore = {
      listThreads: vi.fn(async () => [
        { threadId: "thread-jsonl", preview: "jsonl", modelProvider: "test", createdAt: 1, updatedAt: 1 },
      ]),
    };
    const stateRuntimeStore = {
      listThreadSnapshots: vi.fn(async () => [
        { threadId: "thread-sqlite", preview: "sqlite", modelProvider: "test", createdAt: 2, updatedAt: 2 },
      ]),
    };

    await expect(
      listThreadsForIpc({
        getSessionStoreInstance: () => sessionStore as any,
        getStateRuntimeStoreInstance: async () => stateRuntimeStore as any,
      }),
    ).resolves.toEqual([
      { threadId: "thread-sqlite", preview: "sqlite", modelProvider: "test", createdAt: 2, updatedAt: 2 },
    ]);
    expect(stateRuntimeStore.listThreadSnapshots).toHaveBeenCalledTimes(1);
    expect(sessionStore.listThreads).not.toHaveBeenCalled();
  });

  it("falls back to JSONL session scanning when sqlite listing fails", async () => {
    const sessionStore = {
      listThreads: vi.fn(async () => [
        { threadId: "thread-jsonl", preview: "jsonl", modelProvider: "test", createdAt: 1, updatedAt: 1 },
      ]),
    };

    await expect(
      listThreadsForIpc({
        getSessionStoreInstance: () => sessionStore as any,
        getStateRuntimeStoreInstance: async () => {
          throw new Error("db unavailable");
        },
      }),
    ).resolves.toEqual([
      { threadId: "thread-jsonl", preview: "jsonl", modelProvider: "test", createdAt: 1, updatedAt: 1 },
    ]);
    expect(sessionStore.listThreads).toHaveBeenCalledTimes(1);
  });
});

describe("enqueueTurnForIpc", () => {
  it("queues input immediately when the agent is running", async () => {
    const agent = {
      getIsRunning: vi.fn(() => true),
      enqueueTurn: vi.fn(() => ({ queued: true, queueSize: 1 })),
      runTurn: vi.fn(),
      getThread: vi.fn(),
    };
    const callbacks = { onEvent: vi.fn() };

    await expect(
      enqueueTurnForIpc(agent as any, { content: "等一下，用 Sheet2" }, callbacks)
    ).resolves.toEqual({ success: true, queued: true, queueSize: 1 });

    expect(agent.enqueueTurn).toHaveBeenCalledWith({ content: "等一下，用 Sheet2" }, callbacks);
    expect(agent.runTurn).not.toHaveBeenCalled();
  });

  it("runs immediately when the agent is idle", async () => {
    const agent = {
      getIsRunning: vi.fn(() => false),
      enqueueTurn: vi.fn(),
      runTurn: vi.fn(async () => ({ turnId: "turn-1" })),
      getThread: vi.fn(() => ({ metadata: { threadId: "thread-1" } })),
    };
    const callbacks = { onEvent: vi.fn() };

    await expect(
      enqueueTurnForIpc(agent as any, { content: "现在发送" }, callbacks)
    ).resolves.toEqual({
      success: true,
      queued: false,
      turnId: "turn-1",
      threadId: "thread-1",
    });

    expect(agent.runTurn).toHaveBeenCalledWith({ content: "现在发送" }, callbacks);
    expect(agent.enqueueTurn).not.toHaveBeenCalled();
  });
  it("rejects starting an idle thread while another thread is running", async () => {
    const agent = {
      getIsRunning: vi.fn(() => false),
      enqueueTurn: vi.fn(),
      runTurn: vi.fn(),
      getThread: vi.fn(() => ({ metadata: { threadId: "thread-idle" } })),
    };
    const manager = {
      hasRunningLoopOtherThan: vi.fn(() => true),
    };
    const callbacks = { onEvent: vi.fn() };

    await expect(
      enqueueTurnForIpc(agent as any, { content: "start" }, callbacks, manager as any)
    ).rejects.toThrow("当前已有会话正在执行");

    expect(manager.hasRunningLoopOtherThan).toHaveBeenCalledWith("thread-idle");
    expect(agent.runTurn).not.toHaveBeenCalled();
    expect(agent.enqueueTurn).not.toHaveBeenCalled();
  });
});

describe("prepareAgentForStartTurn", () => {
  it("starts and remembers a new loop before the long-running turn begins", async () => {
    let thread: any = null;
    const agent = {
      getThread: vi.fn(() => thread),
      startThread: vi.fn(async () => {
        thread = { metadata: { threadId: "thread-new" } };
        return "thread-new";
      }),
    };
    const manager = {
      hasRunningLoopOtherThan: vi.fn(() => false),
      rememberLoop: vi.fn(),
    };

    await prepareAgentForStartTurn(agent as any, manager as any);

    expect(agent.startThread).toHaveBeenCalledTimes(1);
    expect(manager.rememberLoop).toHaveBeenCalledWith(agent);
  });

  it("remembers an existing loop without restarting its thread", async () => {
    const agent = {
      getThread: vi.fn(() => ({ metadata: { threadId: "thread-existing" } })),
      startThread: vi.fn(),
    };
    const manager = {
      hasRunningLoopOtherThan: vi.fn(() => false),
      rememberLoop: vi.fn(),
    };

    await prepareAgentForStartTurn(agent as any, manager as any);

    expect(agent.startThread).not.toHaveBeenCalled();
    expect(manager.rememberLoop).toHaveBeenCalledWith(agent);
  });

  it("does not start a new thread when another loop is running", async () => {
    const agent = {
      getThread: vi.fn(() => null),
      startThread: vi.fn(),
    };
    const manager = {
      hasRunningLoopOtherThan: vi.fn(() => true),
      rememberLoop: vi.fn(),
    };

    await expect(prepareAgentForStartTurn(agent as any, manager as any)).rejects.toThrow("当前已有会话正在执行");

    expect(agent.startThread).not.toHaveBeenCalled();
    expect(manager.rememberLoop).not.toHaveBeenCalled();
  });
});

describe("prepareNewThreadForIpc", () => {
  it("rejects creating a new thread while any managed loop is running", async () => {
    const manager = {
      hasRunningLoopOtherThan: vi.fn(() => true),
      prepareNewThread: vi.fn(),
    };

    await expect(prepareNewThreadForIpc({
      agentLoopManagerRef: () => manager as any,
      agentLoopRef: () => null,
    }, "D:\\work")).resolves.toEqual({
      success: false,
      error: "当前已有会话正在执行，请等待完成或停止后再新建会话",
    });

    expect(manager.hasRunningLoopOtherThan).toHaveBeenCalledWith(null);
    expect(manager.prepareNewThread).not.toHaveBeenCalled();
  });

  it("prepares a new managed thread when no loop is running", async () => {
    const manager = {
      hasRunningLoopOtherThan: vi.fn(() => false),
      prepareNewThread: vi.fn(),
    };

    await expect(prepareNewThreadForIpc({
      agentLoopManagerRef: () => manager as any,
      agentLoopRef: () => null,
    }, "D:\\work")).resolves.toEqual({ success: true });

    expect(manager.prepareNewThread).toHaveBeenCalledWith("D:\\work");
  });

  it("rejects legacy single-loop new thread while the loop is running", async () => {
    const agent = {
      getIsRunning: vi.fn(() => true),
      resetThread: vi.fn(),
    };

    await expect(prepareNewThreadForIpc({
      agentLoopManagerRef: () => null,
      agentLoopRef: () => agent as any,
    })).resolves.toMatchObject({
      success: false,
      error: "当前已有会话正在执行，请等待完成或停止后再新建会话",
    });

    expect(agent.resetThread).not.toHaveBeenCalled();
  });

  it("resets the legacy single loop when idle", async () => {
    const agent = {
      getIsRunning: vi.fn(() => false),
      resetThread: vi.fn(async () => undefined),
    };

    await expect(prepareNewThreadForIpc({
      agentLoopManagerRef: () => null,
      agentLoopRef: () => agent as any,
    }, "D:\\work")).resolves.toEqual({ success: true });

    expect(agent.resetThread).toHaveBeenCalledWith("D:\\work");
  });
});
