import { afterEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { PAGE_LAYOUT_TOOL_DEFINITIONS } from "../shared/tools/pageLayoutDefinitions";
import { ToolExecutor } from "../shared/tools";
import { MockHostAdapter } from "./mockHost";
import { installPageLayoutExcel } from "./fakes/officeJsPageLayoutFake";

describe("phase44 printArea / titles / fitToOnePage (Office.js ExcelApi 1.9)", () => {
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
    delete (globalThis as { Office?: unknown }).Office;
  });

  it("schema exposes printArea, titles, fitToPages, fitToOnePage, repeat aliases", () => {
    const setDef = PAGE_LAYOUT_TOOL_DEFINITIONS.find((t) => t.name === "sheet.pageLayout.set");
    expect(setDef).toBeTruthy();
    const props = (setDef!.parameters as { properties: Record<string, unknown> }).properties;
    for (const key of [
      "printArea",
      "printTitleRows",
      "printTitleColumns",
      "repeatRows",
      "repeatColumns",
      "fitToPagesWide",
      "fitToPagesTall",
      "fitToOnePageWide",
      "fitToOnePageTall",
    ]) {
      expect(props[key], key).toBeTruthy();
    }
    expect((props.printArea as { minLength?: number }).minLength).toBe(1);
  });

  it("Office.js set/get printArea + print titles with host readback", async () => {
    const gates = installPageLayoutExcel();
    const adapter = new OfficeJsAdapter();
    const set = await adapter.setSheetPageLayout({
      sheetName: "Sheet1",
      printArea: "A1:H50",
      printTitleRows: "$1:$2",
      printTitleColumns: "$A:$B",
    });
    expect(set.ok).toBe(true);
    if (set.ok) {
      expect(set.data.printArea).toBe("A1:H50");
      expect(set.data.printTitleRows).toBe("$1:$2");
      expect(set.data.printTitleColumns).toBe("$A:$B");
    }
    expect(gates.getCommitted("Sheet1")?.printArea).toBe("A1:H50");
    expect(gates.getCommitted("Sheet1")?.printTitleRows).toBe("$1:$2");
    expect(gates.getCommitted("Sheet1")?.printTitleColumns).toBe("$A:$B");

    const got = await adapter.getSheetPageLayout("Sheet1");
    expect(got.ok).toBe(true);
    if (got.ok) {
      expect(got.data.printArea).toBe("A1:H50");
      expect(got.data.printTitleRows).toBe("$1:$2");
      expect(got.data.printTitleColumns).toBe("$A:$B");
    }
  });

  it("executor repeatRows/repeatColumns aliases map to print titles", async () => {
    installPageLayoutExcel();
    const executor = new ToolExecutor(new OfficeJsAdapter());
    const set = await executor.execute({
      name: "sheet.pageLayout.set",
      arguments: {
        sheetName: "Sheet1",
        printArea: "$A$1:$D$20",
        repeatRows: "$1:$1",
        repeatColumns: "$A:$A",
      },
    });
    expect(set.ok).toBe(true);
    if (set.ok) {
      const data = set.data as {
        printArea: string;
        printTitleRows: string;
        printTitleColumns: string;
      };
      expect(data.printArea).toBe("$A$1:$D$20");
      expect(data.printTitleRows).toBe("$1:$1");
      expect(data.printTitleColumns).toBe("$A:$A");
    }
  });

  it("rejects dual official+alias and empty clear attempts", async () => {
    installPageLayoutExcel();
    const executor = new ToolExecutor(new OfficeJsAdapter());
    for (const args of [
      { sheetName: "Sheet1", printTitleRows: "$1:$1", repeatRows: "$2:$2" },
      { sheetName: "Sheet1", printTitleColumns: "$A:$A", repeatColumns: "$B:$B" },
      { sheetName: "Sheet1", printArea: "" },
      { sheetName: "Sheet1", printTitleRows: "  " },
      { sheetName: "Sheet1", repeatRows: "" },
    ]) {
      const result = await executor.execute({
        name: "sheet.pageLayout.set",
        arguments: args,
      });
      expect(result.ok, JSON.stringify(args)).toBe(false);
    }
  });

  it("fitToOnePageWide/Tall maps to fit pages 1 (desktop parity)", async () => {
    const gates = installPageLayoutExcel();
    const executor = new ToolExecutor(new OfficeJsAdapter());

    const wideOnly = await executor.execute({
      name: "sheet.pageLayout.set",
      arguments: { sheetName: "Sheet1", fitToOnePageWide: true },
    });
    expect(wideOnly.ok).toBe(true);
    if (wideOnly.ok) {
      const data = wideOnly.data as {
        fitToPagesWide: number | null;
        fitToPagesTall: number | null;
        zoomScale: number | null;
      };
      expect(data.fitToPagesWide).toBe(1);
      expect(data.fitToPagesTall).toBeNull();
      expect(data.zoomScale).toBeNull();
    }
    expect(gates.getCommitted("Sheet1")?.fitToPagesWide).toBe(1);

    const both = await executor.execute({
      name: "sheet.pageLayout.set",
      arguments: {
        sheetName: "Sheet1",
        fitToOnePageWide: true,
        fitToOnePageTall: true,
      },
    });
    expect(both.ok).toBe(true);
    if (both.ok) {
      const data = both.data as {
        fitToPagesWide: number | null;
        fitToPagesTall: number | null;
      };
      expect(data.fitToPagesWide).toBe(1);
      expect(data.fitToPagesTall).toBe(1);
    }
  });

  it("fitToOnePage rejects false-only, tall-without-wide, and conflicts", async () => {
    installPageLayoutExcel();
    const executor = new ToolExecutor(new OfficeJsAdapter());
    for (const args of [
      { sheetName: "Sheet1", fitToOnePageWide: false },
      { sheetName: "Sheet1", fitToOnePageTall: true },
      {
        sheetName: "Sheet1",
        fitToOnePageWide: true,
        fitToPagesWide: 2,
      },
      {
        sheetName: "Sheet1",
        fitToOnePageWide: true,
        zoomScale: 100,
      },
    ]) {
      const result = await executor.execute({
        name: "sheet.pageLayout.set",
        arguments: args,
      });
      expect(result.ok, JSON.stringify(args)).toBe(false);
    }
  });

  it("MockHost still honors printArea/titles via executor", async () => {
    const host = new MockHostAdapter();
    const executor = new ToolExecutor(host);
    const set = await executor.execute({
      name: "sheet.pageLayout.set",
      arguments: {
        sheetName: "Sheet1",
        printArea: "B2:C10",
        printTitleRows: "$1:$1",
      },
    });
    expect(set.ok).toBe(true);
    if (set.ok) {
      expect(set.data).toMatchObject({
        printArea: "B2:C10",
        printTitleRows: "$1:$1",
      });
    }
  });
});
