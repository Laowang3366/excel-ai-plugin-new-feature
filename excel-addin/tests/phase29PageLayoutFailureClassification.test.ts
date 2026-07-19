import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { ToolExecutor } from "../shared/tools";
import { installPageLayoutExcel } from "./fakes/officeJsPageLayoutFake";

/**
 * Phase29 failure classification:
 * - ExcelApi 1.9 precheck false/missing/throw → typed unsupported, Excel.run 0, writes 0
 * - missing Excel.run after precheck → typed unsupported, Excel.run 0, writes 0
 * - post-run missing members / empty host values → ordinary FailResult with capability+host
 */
describe("phase29 pageLayout failure classification", () => {
  describe("Office.js adapter", () => {
    let gates: ReturnType<typeof installPageLayoutExcel>;

    beforeEach(() => {
      gates = installPageLayoutExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
    });

    it("missing Excel.run after precheck is typed unsupported, Excel.run 0, writes 0", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      const adapter = new OfficeJsAdapter();
      const get = await adapter.getSheetPageLayout("Sheet1");
      expect(get.ok).toBe(false);
      if (!get.ok) {
        expect(get.unsupported).toBe(true);
        expect(get.capability).toBe("sheet.pageLayout.get");
        expect(get.host).toBe("office-js");
        expect(get.reason).toMatch(/Excel\.run is not available/i);
        expect(get.evidence).toMatch(/Excel\.run/i);
      }
      const set = await adapter.setSheetPageLayout({ sheetName: "Sheet1", paperSize: "a4" });
      expect(set.ok).toBe(false);
      if (!set.ok) {
        expect(set.unsupported).toBe(true);
        expect(set.capability).toBe("sheet.pageLayout.set");
        expect(set.host).toBe("office-js");
        expect(set.reason).toMatch(/Excel\.run is not available/i);
      }
      expect(gates.getExcelRunCalls()).toBe(0);
      expect(gates.getPageLayoutWriteCalls()).toBe(0);
    });

    it("ExcelApi 1.9 false: get/set unsupported with Excel.run 0 and writes 0", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      gates = installPageLayoutExcel({ excelApi19: false });
      const adapter = new OfficeJsAdapter();
      const get = await adapter.getSheetPageLayout("Sheet1");
      expect(get.ok).toBe(false);
      if (!get.ok) {
        expect(get.unsupported).toBe(true);
        expect(get.capability).toBe("sheet.pageLayout.get");
        expect(get.host).toBe("office-js");
      }
      const set = await adapter.setSheetPageLayout({ sheetName: "Sheet1", paperSize: "a4" });
      expect(set.ok).toBe(false);
      if (!set.ok) {
        expect(set.unsupported).toBe(true);
        expect(set.capability).toBe("sheet.pageLayout.set");
        expect(set.host).toBe("office-js");
      }
      expect(gates.getExcelRunCalls()).toBe(0);
      expect(gates.getPageLayoutWriteCalls()).toBe(0);
    });

    it("missing isSetSupported: unsupported, Excel.run 0, writes 0", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      gates = installPageLayoutExcel({ missingIsSetSupported: true });
      const adapter = new OfficeJsAdapter();
      const get = await adapter.getSheetPageLayout("Sheet1");
      expect(get.ok).toBe(false);
      if (!get.ok) expect(get.unsupported).toBe(true);
      expect(gates.getExcelRunCalls()).toBe(0);
      expect(gates.getPageLayoutWriteCalls()).toBe(0);
    });

    it("isSetSupported throws: unsupported, Excel.run 0, writes 0", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      gates = installPageLayoutExcel({ isSetSupportedThrows: true });
      const adapter = new OfficeJsAdapter();
      const get = await adapter.getSheetPageLayout("Sheet1");
      expect(get.ok).toBe(false);
      if (!get.ok) expect(get.unsupported).toBe(true);
      const set = await adapter.setSheetPageLayout({ sheetName: "Sheet1", zoomScale: 100 });
      expect(set.ok).toBe(false);
      if (!set.ok) expect(set.unsupported).toBe(true);
      expect(gates.getExcelRunCalls()).toBe(0);
      expect(gates.getPageLayoutWriteCalls()).toBe(0);
    });

    it("missing paperSize member after precheck is ordinary failure (not typed unsupported)", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      gates = installPageLayoutExcel({ hasPaperSize: false });
      const adapter = new OfficeJsAdapter();
      const get = await adapter.getSheetPageLayout("Sheet1");
      expect(get.ok).toBe(false);
      if (!get.ok) {
        expect(get.reason).toMatch(/paperSize is missing/i);
        expect(get.reason).not.toMatch(/is not supported in this host|isSetSupported/i);
        expect(get.unsupported).not.toBe(true);
        expect(get.capability).toBe("sheet.pageLayout.get");
        expect(get.host).toBe("office-js");
      }
      expect(gates.getExcelRunCalls()).toBe(1);
      expect(gates.getPageLayoutWriteCalls()).toBe(0);
    });

    it("missing zoom member after precheck is ordinary failure (not typed unsupported)", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      gates = installPageLayoutExcel({ hasZoom: false });
      const adapter = new OfficeJsAdapter();
      const get = await adapter.getSheetPageLayout("Sheet1");
      expect(get.ok).toBe(false);
      if (!get.ok) {
        expect(get.reason).toMatch(/zoom is missing/i);
        expect(get.reason).not.toMatch(/is not supported in this host|isSetSupported/i);
        expect(get.unsupported).not.toBe(true);
        expect(get.capability).toBe("sheet.pageLayout.get");
        expect(get.host).toBe("office-js");
      }
      expect(gates.getExcelRunCalls()).toBe(1);
      expect(gates.getPageLayoutWriteCalls()).toBe(0);

      const set = await adapter.setSheetPageLayout({ sheetName: "Sheet1", zoomScale: 120 });
      expect(set.ok).toBe(false);
      if (!set.ok) {
        expect(set.reason).toMatch(/zoom is missing/i);
        expect(set.unsupported).not.toBe(true);
        expect(set.capability).toBe("sheet.pageLayout.set");
        expect(set.host).toBe("office-js");
      }
      expect(gates.getExcelRunCalls()).toBe(2);
      expect(gates.getPageLayoutWriteCalls()).toBe(0);
    });

    it("empty host Worksheet.name is ordinary failure without input echo", async () => {
      gates.setHostSheetName("Sheet1", "");
      const adapter = new OfficeJsAdapter();
      const got = await adapter.getSheetPageLayout("Sheet1");
      expect(got.ok).toBe(false);
      if (!got.ok) {
        expect(got.reason).toMatch(/Worksheet\.name|non-empty string/i);
        expect(got.reason).not.toMatch(/Sheet1/);
        expect(got.unsupported).not.toBe(true);
        expect(got.capability).toBe("sheet.pageLayout.get");
        expect(got.host).toBe("office-js");
      }
      expect(gates.getExcelRunCalls()).toBe(1);
      expect(gates.getPageLayoutWriteCalls()).toBe(0);
    });

    it("non-string host Worksheet.name is ordinary failure without input echo", async () => {
      gates.setHostSheetName("Sheet1", 42);
      const adapter = new OfficeJsAdapter();
      const got = await adapter.getSheetPageLayout("Sheet1");
      expect(got.ok).toBe(false);
      if (!got.ok) {
        expect(got.reason).toMatch(/Worksheet\.name|non-empty string/i);
        expect(got.reason).not.toMatch(/Sheet1/);
        expect(got.unsupported).not.toBe(true);
        expect(got.capability).toBe("sheet.pageLayout.get");
        expect(got.host).toBe("office-js");
      }
      expect(gates.getExcelRunCalls()).toBe(1);
      expect(gates.getPageLayoutWriteCalls()).toBe(0);
    });
  });

  describe("ToolExecutor classification", () => {
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
    });

    it("ordinary: missing paperSize keeps structured host result in detail", async () => {
      const gates = installPageLayoutExcel({ hasPaperSize: false });
      const executor = new ToolExecutor(new OfficeJsAdapter());
      const ordinary = await executor.execute({
        name: "sheet.pageLayout.get",
        arguments: { sheetName: "Sheet1" },
      });
      expect(ordinary.ok).toBe(false);
      if (!ordinary.ok) {
        expect(ordinary.error).toMatch(/paperSize is missing/i);
        expect(ordinary.unsupported).not.toBe(true);
        const detail = ordinary.detail as {
          ok: false;
          unsupported?: boolean;
          capability?: string;
          host?: string;
          reason?: string;
        };
        expect(detail.ok).toBe(false);
        expect(detail.unsupported).not.toBe(true);
        expect(detail.capability).toBe("sheet.pageLayout.get");
        expect(detail.host).toBe("office-js");
        expect(detail.reason).toMatch(/paperSize is missing/i);
      }
      expect(gates.getExcelRunCalls()).toBe(1);
      expect(gates.getPageLayoutWriteCalls()).toBe(0);
    });

    it("ordinary: missing zoom keeps structured host result in detail", async () => {
      const gates = installPageLayoutExcel({ hasZoom: false });
      const executor = new ToolExecutor(new OfficeJsAdapter());
      const ordinary = await executor.execute({
        name: "sheet.pageLayout.get",
        arguments: { sheetName: "Sheet1" },
      });
      expect(ordinary.ok).toBe(false);
      if (!ordinary.ok) {
        expect(ordinary.error).toMatch(/zoom is missing/i);
        expect(ordinary.unsupported).not.toBe(true);
        const detail = ordinary.detail as {
          ok: false;
          unsupported?: boolean;
          capability?: string;
          host?: string;
          reason?: string;
        };
        expect(detail.unsupported).not.toBe(true);
        expect(detail.capability).toBe("sheet.pageLayout.get");
        expect(detail.host).toBe("office-js");
        expect(detail.reason).toMatch(/zoom is missing/i);
      }
      expect(gates.getExcelRunCalls()).toBe(1);
      expect(gates.getPageLayoutWriteCalls()).toBe(0);
    });

    it("ordinary: empty host Worksheet.name keeps structured host result, no input echo", async () => {
      const gates = installPageLayoutExcel();
      gates.setHostSheetName("Sheet1", "");
      const executor = new ToolExecutor(new OfficeJsAdapter());
      const ordinary = await executor.execute({
        name: "sheet.pageLayout.get",
        arguments: { sheetName: "Sheet1" },
      });
      expect(ordinary.ok).toBe(false);
      if (!ordinary.ok) {
        expect(ordinary.error).toMatch(/Worksheet\.name|non-empty string/i);
        expect(ordinary.error).not.toMatch(/Sheet1/);
        expect(ordinary.unsupported).not.toBe(true);
        const detail = ordinary.detail as {
          ok: false;
          unsupported?: boolean;
          capability?: string;
          host?: string;
          reason?: string;
        };
        expect(detail.unsupported).not.toBe(true);
        expect(detail.capability).toBe("sheet.pageLayout.get");
        expect(detail.host).toBe("office-js");
        expect(detail.reason).not.toMatch(/Sheet1/);
      }
      expect(gates.getExcelRunCalls()).toBe(1);
      expect(gates.getPageLayoutWriteCalls()).toBe(0);
    });

    it("ordinary: non-string host Worksheet.name keeps structured host result, no input echo", async () => {
      const gates = installPageLayoutExcel();
      gates.setHostSheetName("Sheet1", { bad: true });
      const executor = new ToolExecutor(new OfficeJsAdapter());
      const ordinary = await executor.execute({
        name: "sheet.pageLayout.get",
        arguments: { sheetName: "Sheet1" },
      });
      expect(ordinary.ok).toBe(false);
      if (!ordinary.ok) {
        expect(ordinary.error).toMatch(/Worksheet\.name|non-empty string/i);
        expect(ordinary.error).not.toMatch(/Sheet1/);
        expect(ordinary.unsupported).not.toBe(true);
        const detail = ordinary.detail as {
          capability?: string;
          host?: string;
          unsupported?: boolean;
        };
        expect(detail.unsupported).not.toBe(true);
        expect(detail.capability).toBe("sheet.pageLayout.get");
        expect(detail.host).toBe("office-js");
      }
      expect(gates.getExcelRunCalls()).toBe(1);
      expect(gates.getPageLayoutWriteCalls()).toBe(0);
    });

    it("typed unsupported: ExcelApi 1.9 false, run 0, writes 0, detail preserves host result", async () => {
      const gates = installPageLayoutExcel({ excelApi19: false });
      const executor = new ToolExecutor(new OfficeJsAdapter());
      const typed = await executor.execute({
        name: "sheet.pageLayout.set",
        arguments: { sheetName: "Sheet1", paperSize: "a4" },
      });
      expect(typed.ok).toBe(false);
      if (!typed.ok) {
        expect(typed.unsupported).toBe(true);
        expect(typed.error).toMatch(/isSetSupported|1\.9/);
        const detail = typed.detail as {
          ok: false;
          unsupported?: boolean;
          capability?: string;
          host?: string;
        };
        expect(detail.unsupported).toBe(true);
        expect(detail.capability).toBe("sheet.pageLayout.set");
        expect(detail.host).toBe("office-js");
      }
      expect(gates.getExcelRunCalls()).toBe(0);
      expect(gates.getPageLayoutWriteCalls()).toBe(0);
    });

    it("typed unsupported: missing Excel.run after precheck, run 0, writes 0", async () => {
      const gates = installPageLayoutExcel();
      delete (globalThis as { Excel?: unknown }).Excel;
      const executor = new ToolExecutor(new OfficeJsAdapter());
      const get = await executor.execute({
        name: "sheet.pageLayout.get",
        arguments: { sheetName: "Sheet1" },
      });
      expect(get.ok).toBe(false);
      if (!get.ok) {
        expect(get.unsupported).toBe(true);
        expect(get.error).toMatch(/Excel\.run is not available/i);
        const detail = get.detail as {
          unsupported?: boolean;
          capability?: string;
          host?: string;
          evidence?: string;
        };
        expect(detail.unsupported).toBe(true);
        expect(detail.capability).toBe("sheet.pageLayout.get");
        expect(detail.host).toBe("office-js");
        expect(detail.evidence).toMatch(/Excel\.run/i);
      }
      const set = await executor.execute({
        name: "sheet.pageLayout.set",
        arguments: { sheetName: "Sheet1", paperSize: "a4" },
      });
      expect(set.ok).toBe(false);
      if (!set.ok) {
        expect(set.unsupported).toBe(true);
        const detail = set.detail as { capability?: string; host?: string };
        expect(detail.capability).toBe("sheet.pageLayout.set");
        expect(detail.host).toBe("office-js");
      }
      expect(gates.getExcelRunCalls()).toBe(0);
      expect(gates.getPageLayoutWriteCalls()).toBe(0);
    });
  });
});
