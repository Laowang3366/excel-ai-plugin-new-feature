import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { ToolExecutor } from "../shared/tools";
import { installInspectExcel } from "./fakes/officeJsInspectFake";
import { MockHostAdapter } from "./mockHost";

describe("phase10 workbook.inspect sheet dimensions", () => {
  describe("Office.js", () => {
    beforeEach(() => {
      installInspectExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
    });

    it("returns per-sheet usedRangeAddress/rowCount/columnCount and empty null/0/0", async () => {
      const adapter = new OfficeJsAdapter();
      const result = await adapter.inspectWorkbook();
      expect(result.ok).toBe(true);
      if (!result.ok) return;

      expect(result.data.workbookName).toBe("Book1.xlsx");
      expect(result.data.activeSheetName).toBe("Sheet1");
      expect(result.data.usedRangeAddress).toBe("Sheet1!A1:C3");
      expect(result.data.sheetCount).toBe(3);

      const sheet1 = result.data.sheets.find((s) => s.name === "Sheet1");
      const empty = result.data.sheets.find((s) => s.name === "Empty");
      const data = result.data.sheets.find((s) => s.name === "Data");

      expect(sheet1).toMatchObject({
        isActive: true,
        usedRangeAddress: "Sheet1!A1:C3",
        rowCount: 3,
        columnCount: 3,
      });
      expect(empty).toMatchObject({
        isActive: false,
        usedRangeAddress: null,
        rowCount: 0,
        columnCount: 0,
      });
      expect(data).toMatchObject({
        usedRangeAddress: "Data!B2:D10",
        rowCount: 9,
        columnCount: 3,
      });
    });
  });

  it("sheet.list does not attach used-range dimensions", async () => {
    const host = new MockHostAdapter();
    const listed = await host.listSheets();
    expect(listed.ok).toBe(true);
    if (listed.ok) {
      for (const sheet of listed.data) {
        expect(sheet.rowCount).toBeUndefined();
        expect(sheet.columnCount).toBeUndefined();
        expect(sheet.usedRangeAddress).toBeUndefined();
      }
    }
  });

  it("executor workbook.inspect remains compatible", async () => {
    const executor = new ToolExecutor(new MockHostAdapter());
    const result = await executor.execute({ name: "workbook.inspect", arguments: {} });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const data = result.data as {
        usedRangeAddress: string | null;
        sheets: Array<{ name: string }>;
      };
      expect(data.usedRangeAddress).toBeTruthy();
      expect(data.sheets.length).toBeGreaterThan(0);
    }
  });

  describe("WPS", () => {
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

    it("keeps address and leaves sheet dimensions unset", async () => {
      const adapter = new WpsJsaAdapter();
      const inspect = await adapter.inspectWorkbook();
      expect(inspect.ok).toBe(true);
      if (!inspect.ok) return;
      expect(inspect.data.usedRangeAddress).toBe("A1:B2");
      for (const sheet of inspect.data.sheets) {
        expect(sheet.rowCount).toBeUndefined();
        expect(sheet.columnCount).toBeUndefined();
        expect(sheet.usedRangeAddress).toBeUndefined();
      }
    });
  });
});
