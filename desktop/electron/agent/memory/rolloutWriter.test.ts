import { describe, expect, it, vi } from "vitest";

import { AsyncRolloutWriter, type RolloutWriteBatch } from "./rolloutWriter";

describe("AsyncRolloutWriter", () => {
  it("queues writes and flushes them as one batch per file", async () => {
    const scheduled: Array<() => void> = [];
    const writeBatch = vi.fn(async () => {});
    const writer = new AsyncRolloutWriter({
      writeBatch,
      scheduleDrain: (run) => scheduled.push(run),
    });

    await writer.enqueue("a.jsonl", ["line-1\n"]);
    await writer.enqueue("a.jsonl", ["line-2\n"]);
    await writer.enqueue("b.jsonl", ["line-3\n"]);

    expect(writeBatch).not.toHaveBeenCalled();

    await writer.flush();

    expect(writeBatch).toHaveBeenCalledTimes(2);
    expect(writeBatch).toHaveBeenCalledWith("a.jsonl", "line-1\nline-2\n");
    expect(writeBatch).toHaveBeenCalledWith("b.jsonl", "line-3\n");
  });

  it("keeps writes enqueued during a flush for the next drain loop", async () => {
    let enqueuedDuringWrite = false;
    const writeBatch = vi.fn(async (filePath: string) => {
      if (filePath === "a.jsonl" && !enqueuedDuringWrite) {
        enqueuedDuringWrite = true;
        await writer.enqueue("a.jsonl", ["line-2\n"]);
      }
    });
    const writer = new AsyncRolloutWriter({ writeBatch, scheduleDrain: () => {} });

    await writer.enqueue("a.jsonl", ["line-1\n"]);

    await writer.flush();

    expect(writeBatch).toHaveBeenCalledTimes(2);
    expect(writeBatch).toHaveBeenNthCalledWith(1, "a.jsonl", "line-1\n");
    expect(writeBatch).toHaveBeenNthCalledWith(2, "a.jsonl", "line-2\n");
  });

  it("surfaces write failures on flush", async () => {
    const writer = new AsyncRolloutWriter({
      writeBatch: vi.fn(async () => {
        throw new Error("disk full");
      }),
      scheduleDrain: () => {},
    });

    await writer.enqueue("a.jsonl", ["line-1\n"]);

    await expect(writer.flush()).rejects.toThrow("disk full");
  });

  it("applies backpressure when the queued line capacity is full", async () => {
    const scheduled: Array<() => void> = [];
    let releaseWrite!: () => void;
    let writeBatch!: RolloutWriteBatch;
    let blockedFirstWrite = false;
    const firstWriteStarted = new Promise<void>((resolve) => {
      writeBatch = vi.fn(async () => {
        if (blockedFirstWrite) return;
        blockedFirstWrite = true;
        resolve();
        await new Promise<void>((release) => {
          releaseWrite = release;
        });
      });
    });
    let writer!: AsyncRolloutWriter;
    writer = new AsyncRolloutWriter({
      maxQueuedLines: 1,
      writeBatch,
      scheduleDrain: (run) => scheduled.push(run),
    });

    await writer.enqueue("a.jsonl", ["line-1\n"]);
    scheduled.shift()?.();
    await firstWriteStarted;

    const secondEnqueue = writer.enqueue("a.jsonl", ["line-2\n"]);
    let secondResolved = false;
    secondEnqueue.then(() => {
      secondResolved = true;
    });
    await Promise.resolve();

    expect(secondResolved).toBe(false);

    releaseWrite();
    await secondEnqueue;
    expect(secondResolved).toBe(true);

    await writer.flush();
  });
});
