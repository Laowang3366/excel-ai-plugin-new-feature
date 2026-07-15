import { describe, expect, it, vi } from "vitest";

import type { IIpcApi } from "./ipcApiTypes";
import { createThreadIpcApi } from "./ipcThreadApi";

describe("createThreadIpcApi", () => {
  it("forwards thread and thread graph calls", async () => {
    const runtimeStatus = vi.fn().mockResolvedValue({ status: "running", idleUnloadMs: 1000 });
    const upsertSpawnEdge = vi.fn().mockResolvedValue({ parentThreadId: "a", childThreadId: "b" });
    const api = createThreadIpcApi(
      () =>
        ({
          thread: { runtimeStatus },
          threadGraph: { upsertSpawnEdge },
        }) as unknown as IIpcApi,
    );

    await expect(api.thread.runtimeStatus()).resolves.toEqual({
      status: "running",
      idleUnloadMs: 1000,
    });
    await expect(api.threadGraph.upsertSpawnEdge("a", "b", "demo")).resolves.toEqual({
      parentThreadId: "a",
      childThreadId: "b",
    });
    expect(upsertSpawnEdge).toHaveBeenCalledWith("a", "b", "demo");
  });

  it("returns safe fallbacks when raw IPC is unavailable", async () => {
    const api = createThreadIpcApi(() => null);

    await expect(api.thread.list()).resolves.toEqual([]);
    await expect(api.thread.runtimeStatus()).resolves.toEqual({
      status: "not_loaded",
      idleUnloadMs: 0,
    });
    await expect(api.threadGraph.closeSpawnEdge("a", "b")).resolves.toBeNull();
    await expect(api.threadGraph.listDescendants("a")).resolves.toEqual([]);
  });
});
