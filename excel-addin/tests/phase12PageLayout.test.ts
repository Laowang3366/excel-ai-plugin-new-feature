import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { ToolExecutor, TOOL_DEFINITIONS } from "../shared/tools";
import { installPageLayoutExcel } from "./fakes/officeJsPageLayoutFake";
import { MockHostAdapter } from "./mockHost";

describe("phase12 sheet.pageLayout", () => {
  it("registers pageLayout tools", () => {
    const names = TOOL_DEFINITIONS.map((tool) => tool.name);
    expect(names).toContain("sheet.pageLayout.get");
    expect(names).toContain("sheet.pageLayout.set");
  });

  describe("Office.js", () => {
    let gates: ReturnType<typeof installPageLayoutExcel>;

    beforeEach(() => {
      gates = installPageLayoutExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
    });

    it("reads defaults and writebacks confirmed fields", async () => {
      const adapter = new OfficeJsAdapter();
      const got = await adapter.getSheetPageLayout("Sheet1");
      expect(got.ok).toBe(true);
      if (got.ok) {
        expect(got.data.orientation).toBe("portrait");
        expect(got.data.printArea).toBeNull();
        expect(got.data.zoomScale).toBe(100);
        expect(got.data.paperSize).toBe("letter");
        expect(got.data.fitToPagesWide).toBeNull();
        expect(got.data.fitToPagesTall).toBeNull();
      }

      const set = await adapter.setSheetPageLayout({
        sheetName: "Sheet1",
        orientation: "landscape",
        centerHorizontally: true,
        centerVertically: true,
        printGridlines: true,
        printHeadings: true,
        margins: { top: 36, bottom: 40, left: 48, right: 44 },
        zoomScale: 120,
        printArea: "A1:D20",
        printTitleRows: "1:1",
        printTitleColumns: "A:A",
      });
      expect(set.ok).toBe(true);
      if (set.ok) {
        expect(set.data.orientation).toBe("landscape");
        expect(set.data.centerHorizontally).toBe(true);
        expect(set.data.centerVertically).toBe(true);
        expect(set.data.printGridlines).toBe(true);
        expect(set.data.printHeadings).toBe(true);
        expect(set.data.margins.top).toBe(36);
        expect(set.data.margins.bottom).toBe(40);
        expect(set.data.margins.left).toBe(48);
        expect(set.data.margins.right).toBe(44);
        expect(set.data.zoomScale).toBe(120);
        expect(set.data.fitToPagesWide).toBeNull();
        expect(set.data.fitToPagesTall).toBeNull();
        expect(set.data.printArea).toBe("A1:D20");
        expect(set.data.printTitleRows).toBe("1:1");
        expect(set.data.printTitleColumns).toBe("A:A");
      }

      const other = await adapter.setSheetPageLayout({
        sheetName: "Sheet2",
        blackAndWhite: true,
      });
      expect(other.ok).toBe(true);
      if (other.ok) {
        expect(other.data.sheetName).toBe("Sheet2");
        expect(other.data.blackAndWhite).toBe(true);
      }
    });

    it("reads null zoomScale when host scale is null (fit-to-pages)", async () => {
      gates.setCommittedZoomScale("Sheet1", null);
      gates.setCommittedFit("Sheet1", 1, 2);
      const adapter = new OfficeJsAdapter();
      const got = await adapter.getSheetPageLayout("Sheet1");
      expect(got.ok).toBe(true);
      if (got.ok) {
        expect(got.data.zoomScale).toBeNull();
        expect(got.data.fitToPagesWide).toBe(1);
        expect(got.data.fitToPagesTall).toBe(2);
        expect(got.data.printArea).toBeNull();
      }
    });

    it("pageLayout stays stale until sync (proves writeback depends on sync)", async () => {
      const Excel = (
        globalThis as unknown as {
          Excel: {
            run: <T>(
              fn: (ctx: {
                workbook: { worksheets: { getItem: (n: string) => unknown } };
                sync: () => Promise<void>;
              }) => Promise<T>,
            ) => Promise<T>;
          };
        }
      ).Excel;

      await Excel.run(async (context) => {
        const ws = context.workbook.worksheets.getItem("Sheet1") as {
          pageLayout: {
            orientation: string;
            paperSize: string;
            zoom: {
              scale?: number | null;
              horizontalFitToPages?: number;
              verticalFitToPages?: number;
            };
            setPrintArea: (a: string) => void;
            getPrintAreaOrNullObject: () => {
              isNullObject: boolean;
              address: string;
              load: (p: string) => void;
            };
          };
        };
        ws.pageLayout.orientation = "Landscape";
        ws.pageLayout.paperSize = "A4";
        ws.pageLayout.zoom = { scale: 150 };
        ws.pageLayout.setPrintArea("B2:C3");

        const before = ws.pageLayout.getPrintAreaOrNullObject();
        before.load("address");
        // Intentionally no sync: committed layout must still be defaults.
        expect(before.isNullObject).toBe(true);
        expect(ws.pageLayout.orientation).toBe("Portrait");
        expect(ws.pageLayout.paperSize).toBe("Letter");
        expect(ws.pageLayout.zoom.scale).toBe(100);
        expect(gates.getPending("Sheet1")?.orientation).toBe("Landscape");
        expect(gates.getPending("Sheet1")?.paperSize).toBe("A4");
        expect(gates.getPending("Sheet1")?.zoomScale).toBe(150);
        expect(gates.getPending("Sheet1")?.printArea).toBe("B2:C3");
        expect(gates.getCommitted("Sheet1")?.printArea).toBeNull();

        await context.sync();
        const after = ws.pageLayout.getPrintAreaOrNullObject();
        after.load("address");
        expect(after.isNullObject).toBe(false);
        expect(after.address).toBe("B2:C3");
        expect(ws.pageLayout.orientation).toBe("Landscape");
        expect(ws.pageLayout.paperSize).toBe("A4");
        expect(ws.pageLayout.zoom.scale).toBe(150);
        expect(gates.getCommitted("Sheet1")?.zoomScale).toBe(150);
        expect(gates.getCommitted("Sheet1")?.printArea).toBe("B2:C3");
      });
    });
  });

  it("executor success and validation branches", async () => {
    const executor = new ToolExecutor(new MockHostAdapter());
    expect(
      (
        await executor.execute({
          name: "sheet.pageLayout.get",
          arguments: { sheetName: "Sheet1" },
        })
      ).ok,
    ).toBe(true);

    expect(
      (
        await executor.execute({
          name: "sheet.pageLayout.set",
          arguments: { sheetName: "Sheet1", orientation: "landscape" },
        })
      ).ok,
    ).toBe(true);

    for (const args of [
      { sheetName: "Sheet1" },
      { sheetName: "Sheet1", orientation: "square" },
      { sheetName: "Sheet1", centerHorizontally: null },
      { sheetName: "Sheet1", zoomScale: 5 },
      { sheetName: "Sheet1", margins: { diagonal: 1 } },
      { sheetName: "Sheet1", printArea: "" },
      { sheetName: "Sheet1", printArea: null },
      { sheetName: "Sheet1", unknown: true },
    ]) {
      const result = await executor.execute({
        name: "sheet.pageLayout.set",
        arguments: args as Record<string, unknown>,
      });
      expect(result.ok).toBe(false);
    }
  });

  it("WPS returns unsupported for pageLayout get/set", async () => {
    const executor = new ToolExecutor(new WpsJsaAdapter());
    for (const name of ["sheet.pageLayout.get", "sheet.pageLayout.set"] as const) {
      const result = await executor.execute({
        name,
        arguments:
          name === "sheet.pageLayout.get"
            ? { sheetName: "Sheet1" }
            : { sheetName: "Sheet1", orientation: "portrait" },
      });
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.unsupported).toBe(true);
    }
  });
});
