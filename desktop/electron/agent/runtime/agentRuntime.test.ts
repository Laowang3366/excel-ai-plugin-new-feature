import { describe, expect, it, vi } from "vitest";

import { AgentLoopManager } from "./agentRuntime";

function loop(threadId: string | null, options: { running?: boolean; resume?: () => Promise<boolean> } = {}) {
  let activeThread = threadId ? { metadata: { threadId } as Record<string, unknown> } : null;
  return {
    getThread: vi.fn(() => activeThread),
    getIsRunning: vi.fn(() => options.running ?? false),
    resumeThread: vi.fn(async (target: string) => {
      const resumed = options.resume ? await options.resume() : true;
      if (resumed) activeThread = { metadata: { threadId: target } };
      return resumed;
    }),
    resetThread: vi.fn(async () => { activeThread = null; }),
    interrupt: vi.fn(),
  };
}

describe("AgentLoopManager", () => {
  it("deduplicates concurrent loads for the same unloaded thread", async () => {
    let finishResume!: () => void;
    const created = loop(null, {
      resume: () => new Promise<boolean>((resolve) => {
        finishResume = () => resolve(true);
      }),
    });
    const createLoop = vi.fn(() => created as any);
    const manager = new AgentLoopManager(createLoop, loop(null) as any);

    const first = manager.getLoopForThread("thread-1");
    const second = manager.getLoopForThread("thread-1");
    expect(createLoop).toHaveBeenCalledTimes(1);
    finishResume();

    await expect(Promise.all([first, second])).resolves.toEqual([created, created]);
    expect(created.resumeThread).toHaveBeenCalledTimes(1);
  });

  it("reloads a thread when its previously mapped loop was idle-unloaded", async () => {
    const stale = loop("thread-1");
    const fresh = loop(null);
    const manager = new AgentLoopManager(() => fresh as any, loop(null) as any);
    manager.rememberLoop(stale as any);
    await stale.resetThread();

    await expect(manager.getLoopForThread("thread-1")).resolves.toBe(fresh);
    expect(fresh.resumeThread).toHaveBeenCalledWith("thread-1");
  });

  it("holds an atomic lock across preparation and turn execution", async () => {
    let release!: () => void;
    const firstLoop = loop("thread-1");
    const secondLoop = loop("thread-2");
    const manager = new AgentLoopManager(vi.fn(), loop(null) as any);

    const first = manager.runWithTurnLock(firstLoop as any, () => new Promise<void>((resolve) => { release = resolve; }));
    await expect(manager.runWithTurnLock(secondLoop as any, async () => undefined))
      .rejects.toThrow("当前已有会话正在执行");
    release();
    await first;
    await expect(manager.runWithTurnLock(secondLoop as any, async () => "ok")).resolves.toBe("ok");
  });

  it("refuses to release a running thread and resets an idle thread", async () => {
    const running = loop("thread-running", { running: true });
    const idle = loop("thread-idle");
    const manager = new AgentLoopManager(vi.fn(), loop(null) as any);
    manager.rememberLoop(running as any);
    manager.rememberLoop(idle as any);

    await expect(manager.releaseThread("thread-running")).resolves.toBe(false);
    await expect(manager.releaseThread("thread-idle")).resolves.toBe(true);
    expect(idle.resetThread).toHaveBeenCalledTimes(1);
  });

  it("updates metadata held by a loaded loop", () => {
    const loaded = loop("thread-1");
    const manager = new AgentLoopManager(vi.fn(), loop(null) as any);
    manager.rememberLoop(loaded as any);

    manager.updateLoadedThreadMetadata("thread-1", { name: "新名称", folderId: "folder-new" });

    expect(loaded.getThread()?.metadata).toMatchObject({
      threadId: "thread-1",
      name: "新名称",
      folderId: "folder-new",
    });
  });

  it("prunes idle-unloaded loops from the managed collection", async () => {
    const unloaded = loop("thread-old");
    const primary = loop(null);
    const manager = new AgentLoopManager(vi.fn(), primary as any);
    manager.rememberLoop(unloaded as any);
    await unloaded.resetThread();

    expect(manager.getAllLoops()).toEqual([primary]);
  });
});
