import { afterEach, describe, expect, it, vi } from "vitest";

import { clearIdleThreadUnloadTimer, scheduleIdleThreadUnload } from "./idleThreadUnload";

describe("idleThreadUnload", () => {
  afterEach(() => {
    vi.useRealTimers();
  });

  it("schedules sweep after the remaining idle delay", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(1000);
    const sweepIdleThread = vi.fn().mockResolvedValue(true);

    const timer = scheduleIdleThreadUnload({
      currentTimer: null,
      isRunning: false,
      hasActiveThread: true,
      getStatus: () => ({
        status: "active",
        idleUnloadMs: 200,
        lastActiveAt: 900,
      }),
      sweepIdleThread,
      scheduleAgain: vi.fn(),
    });

    expect(timer).not.toBeNull();
    await vi.advanceTimersByTimeAsync(99);
    expect(sweepIdleThread).not.toHaveBeenCalled();
    await vi.advanceTimersByTimeAsync(1);
    expect(sweepIdleThread).toHaveBeenCalledTimes(1);
  });

  it("clears existing timer and skips scheduling while running", async () => {
    vi.useFakeTimers();
    const sweepIdleThread = vi.fn().mockResolvedValue(true);
    const currentTimer = setTimeout(sweepIdleThread, 10);

    const timer = scheduleIdleThreadUnload({
      currentTimer,
      isRunning: true,
      hasActiveThread: true,
      getStatus: () => ({
        status: "running",
        idleUnloadMs: 200,
        lastActiveAt: 1,
      }),
      sweepIdleThread,
      scheduleAgain: vi.fn(),
    });

    expect(timer).toBeNull();
    await vi.advanceTimersByTimeAsync(10);
    expect(sweepIdleThread).not.toHaveBeenCalled();
  });

  it("clearIdleThreadUnloadTimer always returns null", () => {
    vi.useFakeTimers();
    const timer = setTimeout(() => {}, 10);

    expect(clearIdleThreadUnloadTimer(timer)).toBeNull();
  });
});
