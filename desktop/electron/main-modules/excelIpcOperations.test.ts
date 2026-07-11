import { describe, expect, it, vi } from "vitest";

import { inspectExcelWorkbookForIpc, readExcelRangeForIpc } from "./excelIpcOperations";

describe("Excel IPC operations", () => {
  it("does not disguise a missing bridge as an empty range", async () => {
    await expect(readExcelRangeForIpc(null, "Sheet1", "A1")).rejects.toThrow("Excel 桥接未初始化");
    await expect(inspectExcelWorkbookForIpc(null)).rejects.toThrow("Excel 桥接未初始化");
  });

  it("propagates bridge read failures", async () => {
    const bridge = {
      readRange: vi.fn().mockRejectedValue(new Error("COM disconnected")),
      inspectWorkbook: vi.fn().mockRejectedValue(new Error("workbook unavailable")),
    };

    await expect(readExcelRangeForIpc(bridge as any, "Sheet1", "A1"))
      .rejects.toThrow("COM disconnected");
    await expect(inspectExcelWorkbookForIpc(bridge as any))
      .rejects.toThrow("workbook unavailable");
  });
});
