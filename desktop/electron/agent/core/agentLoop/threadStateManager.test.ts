import { afterEach, describe, expect, it, vi } from "vitest";
import { DEFAULT_THREAD_IDLE_UNLOAD_MS, ThreadStateManager } from "./threadStateManager";

describe("ThreadStateManager", () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("starts unloaded and exposes the configured idle unload threshold", () => {
    const manager = new ThreadStateManager();

    expect(manager.getSnapshot()).toEqual({
      status: "not_loaded",
      threadId: undefined,
      lastActiveAt: undefined,
      unloadedAt: undefined,
      idleUnloadMs: DEFAULT_THREAD_IDLE_UNLOAD_MS,
    });
    expect(manager.shouldUnload()).toBe(false);
  });

  it("tracks active idle time and only unloads after the threshold", () => {
    const manager = new ThreadStateManager({ idleUnloadMs: 100 });

    manager.markLoaded("thread-1", 1000);

    expect(manager.getSnapshot()).toMatchObject({
      status: "active",
      threadId: "thread-1",
      lastActiveAt: 1000,
      unloadedAt: undefined,
    });
    expect(manager.shouldUnload(1099)).toBe(false);
    expect(manager.shouldUnload(1100)).toBe(true);
  });

  it("does not unload while running and resets idle time when marked idle", () => {
    const manager = new ThreadStateManager({ idleUnloadMs: 100 });

    manager.markRunning("thread-1", 1000);
    expect(manager.shouldUnload(2000)).toBe(false);

    manager.markIdle("thread-1", 2000);
    expect(manager.shouldUnload(2099)).toBe(false);
    expect(manager.shouldUnload(2100)).toBe(true);
  });

  it("records unload time and can clear back to not_loaded", () => {
    const manager = new ThreadStateManager({ idleUnloadMs: 100 });

    manager.markUnloaded(900);
    expect(manager.getSnapshot().status).toBe("not_loaded");

    manager.markLoaded("thread-1", 1000);
    manager.markUnloaded(1200);
    expect(manager.getSnapshot()).toMatchObject({
      status: "unloaded",
      threadId: "thread-1",
      lastActiveAt: 1200,
      unloadedAt: 1200,
    });

    manager.clear();
    expect(manager.getSnapshot()).toMatchObject({
      status: "not_loaded",
      threadId: undefined,
      lastActiveAt: undefined,
      unloadedAt: undefined,
    });
  });

  it("disables idle unload when the threshold is zero or negative", () => {
    const zero = new ThreadStateManager({ idleUnloadMs: 0 });
    zero.markLoaded("thread-1", 1000);
    expect(zero.shouldUnload(999999)).toBe(false);

    const negative = new ThreadStateManager({ idleUnloadMs: -1 });
    negative.markLoaded("thread-1", 1000);
    expect(negative.shouldUnload(999999)).toBe(false);
  });
});
