import { describe, expect, it, vi } from "vitest";

import { InputQueue } from "./inputQueue";
import { PendingInterruptQueue } from "./pendingInterruptQueue";
import { TurnState } from "./turnState";
import {
  drainQueuedTurns,
  enqueueQueuedTurn,
  interruptCurrentTurn,
  shouldRescheduleQueueDrain,
} from "./queuedTurns";

const callbacks = { onEvent: vi.fn() };

describe("queuedTurns", () => {
  it("enqueues supplemental turns only when auto drain is enabled", () => {
    const inputQueue = new InputQueue();

    expect(enqueueQueuedTurn({
      autoDrainInputQueue: true,
      inputQueue,
      turnInput: { content: "继续" },
      callbacks,
    })).toEqual({ queued: true, queueSize: 1 });

    expect(() => enqueueQueuedTurn({
      autoDrainInputQueue: false,
      inputQueue,
      turnInput: { content: "继续" },
      callbacks,
    })).toThrow("Agent 正在中断中");
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
    expect(shouldRescheduleQueueDrain({
      autoDrainInputQueue: true,
      isRunning: false,
      queueSize: 1,
    })).toBe(true);
  });
});
