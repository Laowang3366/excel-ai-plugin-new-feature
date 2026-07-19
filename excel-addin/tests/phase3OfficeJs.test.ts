import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { installSyncGatedExcel } from "./fakes/officeJsSyncGated";

describe("phase3 Office.js", () => {
  let gates: ReturnType<typeof installSyncGatedExcel>;

  beforeEach(() => {
    gates = installSyncGatedExcel();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  it("writes multi-cell numberFormat and title.visible on chart create", async () => {
    const adapter = new OfficeJsAdapter();
    const written = await adapter.writeFormat("Sheet1", "A1:B2", {
      fontBold: true,
      numberFormat: "0.0",
      wrapText: true,
    });
    expect(written.ok).toBe(true);
    if (written.ok) {
      expect(written.data.format.numberFormat).toBe("0.0");
      expect(written.data.format.fontBold).toBe(true);
    }

    const chart = await adapter.createChart({
      sheetName: "Sheet1",
      sourceRange: "A1:B2",
      chartType: "column",
      name: "C1",
      title: "Demo",
    });
    expect(chart.ok).toBe(true);
    if (chart.ok) expect(chart.data.title).toBe("Demo");
    expect(gates.getChartTitleVisible("Sheet1", "C1")).toBe(true);
  });

  it("listCharts only sees worksheets.items after load+sync", async () => {
    const before = gates.captureItemsBeforeSync();
    expect(before).toEqual([]);
    const after = await gates.captureItemsAfterSync();
    expect(after.length).toBe(1);

    const adapter = new OfficeJsAdapter();
    await adapter.createChart({
      sheetName: "Sheet1",
      sourceRange: "A1:B2",
      chartType: "line",
      name: "Trend",
      title: "Sales",
    });
    const listed = await adapter.listCharts();
    expect(listed.ok).toBe(true);
    if (listed.ok) {
      expect(listed.data.some((item) => item.name === "Trend")).toBe(true);
    }
  });

  it("supports table lifecycle and workbook inspect", async () => {
    const adapter = new OfficeJsAdapter();
    const table = await adapter.createTable({
      sheetName: "Sheet1",
      address: "A1:C3",
      name: "T1",
      hasHeaders: true,
    });
    expect(table.ok).toBe(true);
    expect((await adapter.listTables("Sheet1")).ok).toBe(true);
    expect((await adapter.deleteTable("Sheet1", "T1")).ok).toBe(true);
    const inspect = await adapter.inspectWorkbook();
    expect(inspect.ok).toBe(true);
    if (inspect.ok) {
      expect(inspect.data.workbookName).toBe("Book1.xlsx");
      expect(inspect.data.usedRangeAddress).toContain("A1");
    }
  });
});
