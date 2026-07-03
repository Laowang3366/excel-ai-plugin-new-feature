import { describe, expect, it } from "vitest";

import { PendingInterruptQueue } from "./pendingInterruptQueue";

describe("PendingInterruptQueue", () => {
  it("keeps interrupt request ids in arrival order until drained", () => {
    const queue = new PendingInterruptQueue();

    queue.push("request-1");
    queue.push("request-2");

    expect(queue.pendingIds()).toEqual(["request-1", "request-2"]);
    expect(queue.drain()).toEqual(["request-1", "request-2"]);
    expect(queue.pendingIds()).toEqual([]);
  });
});
