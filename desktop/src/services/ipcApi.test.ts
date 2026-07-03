import { afterEach, describe, expect, it, vi } from "vitest";

import { ipcApi } from "./ipcApi";

describe("ipcApi wrapper", () => {
  afterEach(() => {
    const testWindow = (globalThis as any).window;
    if (testWindow) Reflect.deleteProperty(testWindow, "electronAPI");
    Reflect.deleteProperty(globalThis as any, "window");
    vi.restoreAllMocks();
  });

  it("forwards thread runtime status through the wrapper", async () => {
    const runtimeStatus = vi.fn().mockResolvedValue({
      status: "running",
      threadId: "thread-1",
      idleUnloadMs: 120000,
    });
    setElectronApi({ thread: { runtimeStatus } });

    await expect(ipcApi.thread.runtimeStatus()).resolves.toEqual({
      status: "running",
      threadId: "thread-1",
      idleUnloadMs: 120000,
    });
    expect(runtimeStatus).toHaveBeenCalledTimes(1);
  });

  it("forwards Excel readRange expand mode through the wrapper", async () => {
    const readRange = vi.fn().mockResolvedValue({
      values: [[1], [2]],
      address: "H2:H3",
      expanded: true,
      expandMode: "spill",
    });
    setElectronApi({ excel: { readRange } });

    await expect(ipcApi.excel.readRange("Sheet1", "H2", "spill")).resolves.toEqual({
      values: [[1], [2]],
      address: "H2:H3",
      expanded: true,
      expandMode: "spill",
    });
    expect(readRange).toHaveBeenCalledWith("Sheet1", "H2", "spill");
  });

  it("forwards thread graph APIs through the wrapper", async () => {
    const edge = {
      parentThreadId: "parent",
      childThreadId: "child",
      status: "open",
      createdAt: 1,
      label: "demo",
    };
    const upsertSpawnEdge = vi.fn().mockResolvedValue(edge);
    const closeSpawnEdge = vi.fn().mockResolvedValue({ ...edge, status: "closed", closedAt: 2 });
    const listDescendants = vi.fn().mockResolvedValue([
      { threadId: "child", parentThreadId: "parent", depth: 1, edge },
    ]);
    setElectronApi({
      threadGraph: {
        upsertSpawnEdge,
        closeSpawnEdge,
        listDescendants,
      },
    });

    await ipcApi.threadGraph.upsertSpawnEdge("parent", "child", "demo");
    await ipcApi.threadGraph.closeSpawnEdge("parent", "child");
    await ipcApi.threadGraph.listDescendants("parent", "open");

    expect(upsertSpawnEdge).toHaveBeenCalledWith("parent", "child", "demo");
    expect(closeSpawnEdge).toHaveBeenCalledWith("parent", "child");
    expect(listDescendants).toHaveBeenCalledWith("parent", "open");
  });

  it("returns safe fallback values when thread IPC is unavailable", async () => {
    await expect(ipcApi.thread.runtimeStatus()).resolves.toEqual({
      status: "not_loaded",
      idleUnloadMs: 0,
    });
    await expect(ipcApi.threadGraph.closeSpawnEdge("parent", "child")).resolves.toBeNull();
    await expect(ipcApi.threadGraph.listDescendants("parent")).resolves.toEqual([]);
  });
});

function setElectronApi(value: Record<string, unknown>): void {
  Object.defineProperty(globalThis as any, "window", {
    configurable: true,
    value: {},
  });
  Object.defineProperty((globalThis as any).window, "electronAPI", {
    configurable: true,
    value,
  });
}
