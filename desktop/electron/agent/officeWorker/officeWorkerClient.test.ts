import { describe, expect, it, vi } from "vitest";
import type { ChildProcessWithoutNullStreams } from "node:child_process";

import { OfficeWorkerClient } from "./officeWorkerClient";

describe("OfficeWorkerClient process generations", () => {
  it("ignores a late exit event from a replaced worker", () => {
    const client = new OfficeWorkerClient() as unknown as {
      process: ChildProcessWithoutNullStreams | null;
      stdoutBuffer: string;
      pending: Map<string, { timer: NodeJS.Timeout; reject: (error: Error) => void }>;
      onWorkerExit: (worker: ChildProcessWithoutNullStreams, error: Error) => void;
    };
    const oldWorker = { killed: true } as ChildProcessWithoutNullStreams;
    const currentWorker = { killed: false } as ChildProcessWithoutNullStreams;
    const reject = vi.fn();
    const timer = setTimeout(() => undefined, 60_000);
    client.process = currentWorker;
    client.stdoutBuffer = "current-generation";
    client.pending.set("2", { timer, reject });

    client.onWorkerExit(oldWorker, new Error("old worker exited"));

    expect(client.process).toBe(currentWorker);
    expect(client.stdoutBuffer).toBe("current-generation");
    expect(client.pending.has("2")).toBe(true);
    expect(reject).not.toHaveBeenCalled();
    clearTimeout(timer);
  });
});
