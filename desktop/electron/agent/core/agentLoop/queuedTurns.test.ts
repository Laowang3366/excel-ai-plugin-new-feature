import { describe, expect, it, vi } from "vitest";

import { InputQueue } from "./inputQueue";
import { PendingInterruptQueue } from "./pendingInterruptQueue";
import { TurnState } from "./turnState";
import {
  drainQueuedTurns,
  drainQueuedTurnsAndReschedule,
  enqueueQueuedTurn,
  interruptCurrentTurn,
  scheduleQueuedTurnsDrain,
  shouldRescheduleQueueDrain,
} from "./queuedTurns";

const callbacks = { onEvent: vi.fn() };

describe("queuedTurns", () => {
  it("enqueues supplemental turns only when auto drain is enabled", () => {
    const inputQueue = new InputQueue();

    expect(
      enqueueQueuedTurn({
        autoDrainInputQueue: true,
        inputQueue,
        turnInput: { content: "继续" },
        callbacks,
      }),
    ).toEqual({ queued: true, queueSize: 1 });

    expect(() =>
      enqueueQueuedTurn({
        autoDrainInputQueue: false,
        inputQueue,
        turnInput: { content: "继续" },
        callbacks,
      }),
    ).toThrow("Agent 正在中断中");
  });

  it("interrupts current turn and clears pending queues", async () => {
    const inputQueue = new InputQueue();
    const pendingInterruptQueue = new PendingInterruptQueue();
    const turnState = new TurnState();
    const controller = new AbortController();
    let resolveTurn: (() => void) | undefined;
    turnState.abortController = controller;
    turnState.turnCompletionPromise = new Promise((resolve) => {
      resolveTurn = resolve;
    });
    inputQueue.enqueue({ input: { content: "排队" }, callbacks });

    const interruptPromise = interruptCurrentTurn({
      requestId: "req-1",
      pendingInterruptQueue,
      inputQueue,
      turnState,
      disableAutoDrain: vi.fn(),
    });
    resolveTurn?.();
    await interruptPromise;

    expect(controller.signal.aborted).toBe(true);
    expect(inputQueue.size()).toBe(0);
    expect(pendingInterruptQueue.pendingIds()).toEqual([]);
  });

  it("drains queued turns until the queue is empty or running resumes", async () => {
    const inputQueue = new InputQueue();
    const runTurn = vi.fn().mockResolvedValue({});
    inputQueue.enqueue({ input: { content: "1" }, callbacks });
    inputQueue.enqueue({ input: { content: "2" }, callbacks });

    await drainQueuedTurns({
      inputQueue,
      isRunning: () => false,
      runTurn,
    });

    expect(runTurn).toHaveBeenCalledTimes(2);
    expect(inputQueue.size()).toBe(0);
    expect(
      shouldRescheduleQueueDrain({
        autoDrainInputQueue: true,
        isRunning: false,
        queueSize: 1,
      }),
    ).toBe(true);
  });

  it("schedules a queue drain only when allowed", async () => {
    const drain = vi.fn().mockResolvedValue(undefined);
    let isDraining = false;

    scheduleQueuedTurnsDrain({
      autoDrainInputQueue: true,
      isDrainingInputQueue: false,
      isRunning: false,
      setDraining: (value) => {
        isDraining = value;
      },
      drain,
    });

    expect(isDraining).toBe(true);
    expect(drain).not.toHaveBeenCalled();
    await Promise.resolve();
    expect(drain).toHaveBeenCalledTimes(1);

    scheduleQueuedTurnsDrain({
      autoDrainInputQueue: false,
      isDrainingInputQueue: false,
      isRunning: false,
      setDraining: vi.fn(),
      drain,
    });
    expect(drain).toHaveBeenCalledTimes(1);
  });

  it("clears draining state and reschedules when queue remains", async () => {
    const inputQueue = new InputQueue();
    inputQueue.enqueue({ input: { content: "1" }, callbacks });
    const runTurn = vi.fn(async () => ({
      turnId: "turn-1",
      threadId: "thread-1",
      status: "completed" as const,
      items: [],
      startedAt: 1,
      completedAt: 2,
    }));
    const setDraining = vi.fn();
    const scheduleDrain = vi.fn();

    await drainQueuedTurnsAndReschedule({
      inputQueue,
      isRunning: () => false,
      autoDrainInputQueue: () => {
        inputQueue.enqueue({ input: { content: "2" }, callbacks });
        return true;
      },
      runTurn,
      setDraining,
      scheduleDrain,
    });

    expect(setDraining).toHaveBeenCalledWith(false);
    expect(scheduleDrain).toHaveBeenCalledTimes(1);
  });
});
