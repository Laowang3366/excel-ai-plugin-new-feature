import { describe, expect, it, vi } from "vitest";

import { InputQueue } from "./inputQueue";
import type { AgentTurnCallbacks, AgentTurnInput } from "../../shared/types";

function input(content: string): AgentTurnInput {
  return { content };
}

function callbacks(): AgentTurnCallbacks {
  return { onEvent: vi.fn() };
}

describe("InputQueue", () => {
  it("dequeues queued inputs in FIFO order", () => {
    const queue = new InputQueue();

    queue.enqueue({ input: input("first"), callbacks: callbacks() });
    queue.enqueue({ input: input("second"), callbacks: callbacks() });

    expect(queue.size()).toBe(2);
    expect(queue.dequeue()?.input.content).toBe("first");
    expect(queue.dequeue()?.input.content).toBe("second");
    expect(queue.dequeue()).toBeUndefined();
    expect(queue.size()).toBe(0);
  });

  it("clears queued inputs when the active turn is interrupted", () => {
    const queue = new InputQueue();

    queue.enqueue({ input: input("补充 Sheet2"), callbacks: callbacks() });
    queue.enqueue({ input: input("再补充筛选条件"), callbacks: callbacks() });

    expect(queue.clear()).toBe(2);
    expect(queue.size()).toBe(0);
    expect(queue.dequeue()).toBeUndefined();
  });

  it("rejects input when the bounded queue is full", () => {
    const queue = new InputQueue(1);

    queue.enqueue({ input: input("first"), callbacks: callbacks() });

    expect(() => queue.enqueue({ input: input("second"), callbacks: callbacks() })).toThrow(
      "运行中输入队列已满",
    );
    expect(queue.size()).toBe(1);
  });
});
