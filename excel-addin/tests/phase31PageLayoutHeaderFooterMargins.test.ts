import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { PAGE_LAYOUT_TOOL_DEFINITIONS } from "../shared/tools/pageLayoutDefinitions";
import { ToolExecutor } from "../shared/tools";
import { installPageLayoutExcel } from "./fakes/officeJsPageLayoutFake";
import { MockHostAdapter } from "./mockHost";

describe("phase31 sheet.pageLayout margins header/footer", () => {
  describe("schema", () => {
    it("set schema margins include header and footer", () => {
      const setDef = PAGE_LAYOUT_TOOL_DEFINITIONS.find((t) => t.name === "sheet.pageLayout.set");
      expect(setDef).toBeDefined();
      const margins = (setDef!.parameters.properties as Record<string, Record<string, unknown>>)
        .margins as { properties?: Record<string, unknown>; additionalProperties?: boolean };
      expect(margins.properties?.header).toEqual({ type: "number" });
      expect(margins.properties?.footer).toEqual({ type: "number" });
      expect(margins.additionalProperties).toBe(false);
    });
  });

  describe("executor validation", () => {
    it("accepts header:0 and positive footer", async () => {
      const executor = new ToolExecutor(new MockHostAdapter());
      const result = await executor.execute({
        name: "sheet.pageLayout.set",
        arguments: { sheetName: "Sheet1", margins: { header: 0, footer: 48 } },
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        const data = result.data as { margins: { header: number; footer: number } };
        expect(data.margins.header).toBe(0);
        expect(data.margins.footer).toBe(48);
      }
    });

    it("rejects invalid header/footer and unknown margin keys", async () => {
      const executor = new ToolExecutor(new MockHostAdapter());
      const cases: Array<Record<string, unknown>> = [
        { sheetName: "Sheet1", margins: { header: -1 } },
        { sheetName: "Sheet1", margins: { footer: Number.NaN } },
        { sheetName: "Sheet1", margins: { header: Number.POSITIVE_INFINITY } },
        { sheetName: "Sheet1", margins: { footer: null } },
        { sheetName: "Sheet1", margins: { header: "0" } },
        { sheetName: "Sheet1", margins: { headerMargin: 12 } },
        { sheetName: "Sheet1", margins: { gutter: 1 } },
      ];
      for (const args of cases) {
        const result = await executor.execute({ name: "sheet.pageLayout.set", arguments: args });
        expect(result.ok).toBe(false);
      }
    });
  });

  describe("Office.js host readback", () => {
    let gates: ReturnType<typeof installPageLayoutExcel>;

    beforeEach(() => {
      gates = installPageLayoutExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
    });

    it("sets header/footer from non-default host values with sync readback", async () => {
      gates.setCommittedHeaderFooterMargins("Sheet1", 12, 18);
      const adapter = new OfficeJsAdapter();
      const before = await adapter.getSheetPageLayout("Sheet1");
      expect(before.ok).toBe(true);
      if (before.ok) {
        expect(before.data.margins.header).toBe(12);
        expect(before.data.margins.footer).toBe(18);
      }

      const set = await adapter.setSheetPageLayout({
        sheetName: "Sheet1",
        margins: { header: 0, footer: 54 },
      });
      expect(set.ok).toBe(true);
      if (set.ok) {
        expect(set.data.margins.header).toBe(0);
        expect(set.data.margins.footer).toBe(54);
      }
      expect(gates.getCommitted("Sheet1")?.headerMargin).toBe(0);
      expect(gates.getCommitted("Sheet1")?.footerMargin).toBe(54);

      const got = await adapter.getSheetPageLayout("Sheet1");
      expect(got.ok).toBe(true);
      if (got.ok) {
        expect(got.data.margins.header).toBe(0);
        expect(got.data.margins.footer).toBe(54);
      }
    });

    it("missing headerMargin/footerMargin is ordinary failure for get and matching set", async () => {
      for (const opts of [{ hasHeaderMargin: false }, { hasFooterMargin: false }] as const) {
        delete (globalThis as { Excel?: unknown }).Excel;
        delete (globalThis as { Office?: unknown }).Office;
        const local = installPageLayoutExcel(opts);
        const adapter = new OfficeJsAdapter();

        const got = await adapter.getSheetPageLayout("Sheet1");
        expect(got.ok).toBe(false);
        if (!got.ok) {
          expect(got.unsupported).not.toBe(true);
          expect(got.capability).toBe("sheet.pageLayout.get");
          expect(got.host).toBe("office-js");
          expect(got.reason).toMatch(/headerMargin|footerMargin|missing/i);
        }

        const setArgs =
          "hasHeaderMargin" in opts && opts.hasHeaderMargin === false
            ? { sheetName: "Sheet1", margins: { header: 10 } }
            : { sheetName: "Sheet1", margins: { footer: 10 } };
        const set = await adapter.setSheetPageLayout(setArgs);
        expect(set.ok).toBe(false);
        if (!set.ok) {
          expect(set.unsupported).not.toBe(true);
          expect(set.capability).toBe("sheet.pageLayout.set");
          expect(set.host).toBe("office-js");
          expect(set.reason).toMatch(/headerMargin|footerMargin|missing/i);
        }
        expect(local.getExcelRunCalls()).toBeGreaterThanOrEqual(1);
      }
    });
  });

  describe("WPS", () => {
    it("still returns typed unsupported for pageLayout get/set", async () => {
      const executor = new ToolExecutor(new WpsJsaAdapter());
      for (const name of ["sheet.pageLayout.get", "sheet.pageLayout.set"] as const) {
        const result = await executor.execute({
          name,
          arguments:
            name === "sheet.pageLayout.get"
              ? { sheetName: "Sheet1" }
              : { sheetName: "Sheet1", margins: { header: 0 } },
        });
        expect(result.ok).toBe(false);
        if (!result.ok) expect(result.unsupported).toBe(true);
      }
    });
  });
});
