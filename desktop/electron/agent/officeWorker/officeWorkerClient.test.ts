import { describe, expect, it, vi } from "vitest";
import type { ChildProcessWithoutNullStreams } from "node:child_process";

const loggerMocks = vi.hoisted(() => ({
  error: vi.fn(),
  warn: vi.fn(),
  info: vi.fn(),
  debug: vi.fn(),
}));

vi.mock("../../shared/logger", () => ({
  createLogger: () => loggerMocks,
}));

import { OfficeWorkerClient, OfficeWorkerError, validateWorkerResult } from "./officeWorkerClient";

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

  it("logs a stable event when the current worker stops", () => {
    loggerMocks.error.mockClear();
    const client = new OfficeWorkerClient() as unknown as {
      process: ChildProcessWithoutNullStreams | null;
      stdoutBuffer: string;
      pending: Map<string, { timer: NodeJS.Timeout; reject: (error: Error) => void }>;
      onWorkerExit: (worker: ChildProcessWithoutNullStreams, error: Error) => void;
    };
    const worker = { killed: false } as ChildProcessWithoutNullStreams;
    client.process = worker;
    client.stdoutBuffer = "buffer";
    client.pending = new Map();

    client.onWorkerExit(worker, new Error("worker exited"));

    expect(loggerMocks.error).toHaveBeenCalledWith(
      "Office Worker stopped",
      expect.objectContaining({
        event: "desktop.office_worker.stopped",
        message: "worker exited",
      }),
    );
  });
});

describe("validateWorkerResult", () => {
  it("accepts the protocol v2 range write result", () => {
    const result = { written: 2, dynamicCells: 1, arrayCells: 0, plainCells: 1 };
    expect(validateWorkerResult("excel.range.write", result)).toBe(result);
  });

  it("rejects an old Worker range write result", () => {
    expect(() => validateWorkerResult("excel.range.write", { written: 2 })).toThrowError(
      expect.objectContaining<Partial<OfficeWorkerError>>({ code: "protocol_invalid_result" }),
    );
  });
});
