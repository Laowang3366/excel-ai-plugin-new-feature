import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";

describe("phase3 WPS unsupported contracts", () => {
  beforeEach(() => {
    (globalThis as unknown as { window: unknown }).window = globalThis;
    (globalThis as unknown as { Application: unknown }).Application = {
      Name: "WPS",
      ActiveWorkbook: {
        Name: "Book1.xlsx",
        ActiveSheet: { Name: "Sheet1", UsedRange: { Address: "A1:B2" }, Range: () => ({}) },
        Worksheets: {
          Count: 1,
          Item: () => ({ Name: "Sheet1", Index: 1 }),
        },
      },
    };
  });
  afterEach(() => {
    delete (globalThis as { Application?: unknown }).Application;
  });

  it("returns typed unsupported for format/table/chart", async () => {
    const adapter = new WpsJsaAdapter();
    for (const result of [
      await adapter.readFormat("Sheet1", "A1"),
      await adapter.writeFormat("Sheet1", "A1", { fontBold: true }),
      await adapter.listTables(),
      await adapter.createTable({ sheetName: "Sheet1", address: "A1:B2" }),
      await adapter.deleteTable("Sheet1", "T1"),
      await adapter.listCharts(),
      await adapter.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2" }),
      await adapter.deleteChart("Sheet1", "C1"),
    ]) {
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(true);
        expect(result.host).toBe("wps-jsa");
      }
    }
  });

  it("inspects workbook without fabricating used range", async () => {
    const adapter = new WpsJsaAdapter();
    const inspect = await adapter.inspectWorkbook();
    expect(inspect.ok).toBe(true);
    if (inspect.ok) {
      expect(inspect.data.workbookName).toBe("Book1.xlsx");
      expect(inspect.data.activeSheetName).toBe("Sheet1");
      expect(inspect.data.usedRangeAddress).toBe("A1:B2");
    }
  });
});
