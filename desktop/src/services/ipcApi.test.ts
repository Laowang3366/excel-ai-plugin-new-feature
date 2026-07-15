import { afterEach, describe, expect, it, vi } from "vitest";

import { ipcApi } from "./ipcApi";

// @MOCK_INTERFACE: tests install partial window.electronAPI shapes that must stay aligned with preload exposure.
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

  it("forwards the fixed Office application enum through the wrapper", async () => {
    const launchOffice = vi.fn().mockResolvedValue({ success: true });
    setElectronApi({ app: { launchOffice } });

    await expect(ipcApi.app.launchOffice("word")).resolves.toEqual({ success: true });
    expect(launchOffice).toHaveBeenCalledWith("word");
  });

  it("forwards privacy export selection and execution through the wrapper", async () => {
    const selectExportPath = vi.fn().mockResolvedValue({
      canceled: false,
      filePaths: ["C:\\exports\\wengge"],
    });
    const exportUserData = vi.fn().mockResolvedValue({
      success: true,
      exportPath: "C:\\exports\\wengge",
    });
    setElectronApi({ app: { selectExportPath, exportUserData } });

    await expect(ipcApi.app.selectExportPath()).resolves.toMatchObject({ canceled: false });
    await expect(ipcApi.app.exportUserData("C:\\exports\\wengge")).resolves.toMatchObject({
      success: true,
    });
    expect(exportUserData).toHaveBeenCalledWith("C:\\exports\\wengge");
  });

  it("forwards explicit local data erasure confirmation through the wrapper", async () => {
    const eraseUserData = vi.fn().mockResolvedValue({
      success: true,
      erasedCategories: ["settings", "sessions"],
      errors: [],
    });
    setElectronApi({ app: { eraseUserData } });

    await expect(
      ipcApi.app.eraseUserData({ confirmation: "ERASE LOCAL DATA" }),
    ).resolves.toMatchObject({ success: true });
    expect(eraseUserData).toHaveBeenCalledWith({ confirmation: "ERASE LOCAL DATA" });
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

  it("forwards Office automation document and transaction operations", async () => {
    const listDocuments = vi.fn().mockResolvedValue({ success: true, data: [] });
    const undo = vi.fn().mockResolvedValue({ success: true, data: { id: "tx", status: "undone" } });
    setElectronApi({
      office: { automation: { documents: { list: listDocuments }, transactions: { undo } } },
    });

    await ipcApi.office.automation.documents.list("excel");
    await ipcApi.office.automation.transactions.undo("tx", true);

    expect(listDocuments).toHaveBeenCalledWith("excel");
    expect(undo).toHaveBeenCalledWith("tx", true);
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
    const listDescendants = vi
      .fn()
      .mockResolvedValue([{ threadId: "child", parentThreadId: "parent", depth: 1, edge }]);
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

  it("forwards batch folder file listing through the wrapper", async () => {
    const listFilesBatch = vi.fn().mockResolvedValue({
      "C:\\docs": [{ fileName: "a.xlsx", filePath: "C:\\docs\\a.xlsx" }],
      "C:\\decks": [{ fileName: "b.pptx", filePath: "C:\\decks\\b.pptx" }],
    });
    setElectronApi({ folder: { listFilesBatch } });

    await expect(ipcApi.folder.listFilesBatch(["C:\\docs", "C:\\decks"])).resolves.toEqual({
      "C:\\docs": [{ fileName: "a.xlsx", filePath: "C:\\docs\\a.xlsx" }],
      "C:\\decks": [{ fileName: "b.pptx", filePath: "C:\\decks\\b.pptx" }],
    });
    expect(listFilesBatch).toHaveBeenCalledWith(["C:\\docs", "C:\\decks"]);
  });

  it("falls back to single folder listings when batch IPC is unavailable", async () => {
    const listFiles = vi
      .fn()
      .mockResolvedValueOnce([{ fileName: "a.xlsx", filePath: "C:\\docs\\a.xlsx" }])
      .mockResolvedValueOnce([{ fileName: "b.pptx", filePath: "C:\\decks\\b.pptx" }]);
    setElectronApi({ folder: { listFiles } });

    await expect(ipcApi.folder.listFilesBatch(["C:\\docs", "C:\\decks"])).resolves.toEqual({
      "C:\\docs": [{ fileName: "a.xlsx", filePath: "C:\\docs\\a.xlsx" }],
      "C:\\decks": [{ fileName: "b.pptx", filePath: "C:\\decks\\b.pptx" }],
    });
    expect(listFiles).toHaveBeenNthCalledWith(1, "C:\\docs");
    expect(listFiles).toHaveBeenNthCalledWith(2, "C:\\decks");
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
