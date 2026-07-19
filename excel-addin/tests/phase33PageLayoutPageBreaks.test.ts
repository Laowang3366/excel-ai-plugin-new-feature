import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { buildAdvancedExcelBoundary } from "../shared/prompts/advancedExcelBoundary";
import { PAGE_LAYOUT_TOOL_DEFINITIONS } from "../shared/tools/pageLayoutDefinitions";
import { ToolExecutor } from "../shared/tools";
import { installPageLayoutExcel } from "./fakes/officeJsPageLayoutFake";
import { MockHostAdapter } from "./mockHost";

describe("phase33 sheet.pageLayout manual page breaks", () => {
  describe("schema", () => {
    it("exposes clearPageBreaks boolean and A1 arrays with bounds", () => {
      const setDef = PAGE_LAYOUT_TOOL_DEFINITIONS.find((t) => t.name === "sheet.pageLayout.set");
      expect(setDef).toBeDefined();
      const props = setDef!.parameters.properties as Record<string, Record<string, unknown>>;
      expect(props.clearPageBreaks?.type).toBe("boolean");
      for (const key of ["horizontalPageBreaks", "verticalPageBreaks"] as const) {
        const arr = props[key] as {
          type?: string;
          maxItems?: number;
          items?: { type?: string; minLength?: number };
        };
        expect(arr.type).toBe("array");
        expect(arr.maxItems).toBe(1000);
        expect(arr.items?.type).toBe("string");
        expect(arr.items?.minLength).toBe(1);
      }
    });
  });

  describe("executor + MockHost", () => {
    it("trims A4, appends, []/false no-op, clear+replace", async () => {
      const executor = new ToolExecutor(new MockHostAdapter());
      const first = await executor.execute({
        name: "sheet.pageLayout.set",
        arguments: {
          sheetName: "Sheet1",
          horizontalPageBreaks: ["  A4  "],
          verticalPageBreaks: ["C1"],
        },
      });
      expect(first.ok).toBe(true);
      if (first.ok) {
        const data = first.data as {
          horizontalPageBreaks: string[];
          verticalPageBreaks: string[];
        };
        expect(data.horizontalPageBreaks).toEqual(["A4"]);
        expect(data.verticalPageBreaks).toEqual(["C1"]);
      }

      const append = await executor.execute({
        name: "sheet.pageLayout.set",
        arguments: { sheetName: "Sheet1", horizontalPageBreaks: ["A10"] },
      });
      expect(append.ok).toBe(true);
      if (append.ok) {
        expect((append.data as { horizontalPageBreaks: string[] }).horizontalPageBreaks).toEqual([
          "A4",
          "A10",
        ]);
      }

      const empty = await executor.execute({
        name: "sheet.pageLayout.set",
        arguments: { sheetName: "Sheet1", horizontalPageBreaks: [] },
      });
      expect(empty.ok).toBe(true);
      if (empty.ok) {
        expect((empty.data as { horizontalPageBreaks: string[] }).horizontalPageBreaks).toEqual([
          "A4",
          "A10",
        ]);
      }

      const noClear = await executor.execute({
        name: "sheet.pageLayout.set",
        arguments: { sheetName: "Sheet1", clearPageBreaks: false },
      });
      expect(noClear.ok).toBe(true);
      if (noClear.ok) {
        expect((noClear.data as { horizontalPageBreaks: string[] }).horizontalPageBreaks).toEqual([
          "A4",
          "A10",
        ]);
      }

      const cleared = await executor.execute({
        name: "sheet.pageLayout.set",
        arguments: {
          sheetName: "Sheet1",
          clearPageBreaks: true,
          horizontalPageBreaks: ["B2"],
          verticalPageBreaks: ["D1"],
        },
      });
      expect(cleared.ok).toBe(true);
      if (cleared.ok) {
        const data = cleared.data as {
          horizontalPageBreaks: string[];
          verticalPageBreaks: string[];
        };
        expect(data.horizontalPageBreaks).toEqual(["B2"]);
        expect(data.verticalPageBreaks).toEqual(["D1"]);
      }
    });

    it("rejects invalid arrays and unknown keys", async () => {
      const executor = new ToolExecutor(new MockHostAdapter());
      const cases: Array<Record<string, unknown>> = [
        { sheetName: "Sheet1", horizontalPageBreaks: null },
        { sheetName: "Sheet1", verticalPageBreaks: "A4" },
        { sheetName: "Sheet1", horizontalPageBreaks: ["  "] },
        { sheetName: "Sheet1", horizontalPageBreaks: [1] },
        {
          sheetName: "Sheet1",
          horizontalPageBreaks: Array.from({ length: 1001 }, () => "A2"),
        },
        { sheetName: "Sheet1", pageBreaks: [] },
      ];
      for (const args of cases) {
        const result = await executor.execute({ name: "sheet.pageLayout.set", arguments: args });
        expect(result.ok).toBe(false);
      }
    });
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

    it("get empty arrays; add/append/clear/replace; [] no-op; host address normalize", async () => {
      const adapter = new OfficeJsAdapter();
      const empty = await adapter.getSheetPageLayout("Sheet1");
      expect(empty.ok).toBe(true);
      if (empty.ok) {
        expect(empty.data.horizontalPageBreaks).toEqual([]);
        expect(empty.data.verticalPageBreaks).toEqual([]);
      }

      const set = await adapter.setSheetPageLayout({
        sheetName: "Sheet1",
        horizontalPageBreaks: ["A4"],
        verticalPageBreaks: ["C1"],
      });
      expect(set.ok).toBe(true);
      if (set.ok) {
        expect(set.data.horizontalPageBreaks).toEqual(["A4"]);
        expect(set.data.verticalPageBreaks).toEqual(["C1"]);
      }
      expect(gates.getCommittedPageBreaks("Sheet1")).toEqual({
        horizontal: ["A4"],
        vertical: ["C1"],
      });

      const append = await adapter.setSheetPageLayout({
        sheetName: "Sheet1",
        horizontalPageBreaks: ["A10"],
      });
      expect(append.ok).toBe(true);
      if (append.ok) expect(append.data.horizontalPageBreaks).toEqual(["A4", "A10"]);

      const noop = await adapter.setSheetPageLayout({
        sheetName: "Sheet1",
        horizontalPageBreaks: [],
      });
      expect(noop.ok).toBe(true);
      if (noop.ok) expect(noop.data.horizontalPageBreaks).toEqual(["A4", "A10"]);

      const replaced = await adapter.setSheetPageLayout({
        sheetName: "Sheet1",
        clearPageBreaks: true,
        horizontalPageBreaks: ["B2"],
      });
      expect(replaced.ok).toBe(true);
      if (replaced.ok) {
        expect(replaced.data.horizontalPageBreaks).toEqual(["B2"]);
        expect(replaced.data.verticalPageBreaks).toEqual([]);
      }

      // host address forms with sheet/$ normalize
      gates.setCommittedPageBreaks("Sheet1", {
        horizontal: ["Sheet1!$A$4"],
        vertical: ["$C$1"],
      });
      const normalized = await adapter.getSheetPageLayout("Sheet1");
      expect(normalized.ok).toBe(true);
      if (normalized.ok) {
        expect(normalized.data.horizontalPageBreaks).toEqual(["A4"]);
        expect(normalized.data.verticalPageBreaks).toEqual(["C1"]);
      }

      // not input echo
      gates.setHostSheetName("Sheet1", "HostPB");
      const again = await adapter.getSheetPageLayout("Sheet1");
      expect(again.ok).toBe(true);
      if (again.ok) expect(again.data.sheetName).toBe("HostPB");
    });

    it("bad host address and missing collection APIs are ordinary failures", async () => {
      const adapter = new OfficeJsAdapter();
      gates.setCommittedPageBreaks("Sheet1", { horizontal: ["A1:B2"] });
      const bad = await adapter.getSheetPageLayout("Sheet1");
      expect(bad.ok).toBe(false);
      if (!bad.ok) {
        expect(bad.unsupported).not.toBe(true);
        expect(bad.capability).toBe("sheet.pageLayout.get");
        expect(bad.host).toBe("office-js");
      }

      for (const opts of [
        { hasHorizontalPageBreaks: false },
        { hasVerticalPageBreaks: false },
        { hasPageBreakItems: false },
      ]) {
        delete (globalThis as { Excel?: unknown }).Excel;
        delete (globalThis as { Office?: unknown }).Office;
        installPageLayoutExcel(opts);
        const local = new OfficeJsAdapter();
        const got = await local.getSheetPageLayout("Sheet1");
        expect(got.ok).toBe(false);
        if (!got.ok) {
          expect(got.unsupported).not.toBe(true);
          expect(got.capability).toBe("sheet.pageLayout.get");
        }
      }

      // Missing getCellAfterBreak only surfaces when a break item exists.
      {
        delete (globalThis as { Excel?: unknown }).Excel;
        delete (globalThis as { Office?: unknown }).Office;
        const localGates = installPageLayoutExcel({ hasGetCellAfterBreak: false });
        localGates.setCommittedPageBreaks("Sheet1", { horizontal: ["A4"] });
        const local = new OfficeJsAdapter();
        const got = await local.getSheetPageLayout("Sheet1");
        expect(got.ok).toBe(false);
        if (!got.ok) {
          expect(got.unsupported).not.toBe(true);
          expect(got.capability).toBe("sheet.pageLayout.get");
        }
      }

      for (const opts of [{ hasPageBreakAdd: false }, { hasPageBreakRemove: false }, { hasHorizontalPageBreaks: false }]) {
        delete (globalThis as { Excel?: unknown }).Excel;
        delete (globalThis as { Office?: unknown }).Office;
        installPageLayoutExcel(opts);
        const local = new OfficeJsAdapter();
        const set = await local.setSheetPageLayout({
          sheetName: "Sheet1",
          clearPageBreaks: "hasPageBreakRemove" in opts ? true : undefined,
          horizontalPageBreaks:
            "hasPageBreakAdd" in opts || "hasHorizontalPageBreaks" in opts ? ["A4"] : undefined,
        });
        expect(set.ok).toBe(false);
        if (!set.ok) {
          expect(set.unsupported).not.toBe(true);
          expect(set.capability).toBe("sheet.pageLayout.set");
        }
      }
    });

    it("ExcelApi 1.9 off remains typed unsupported", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installPageLayoutExcel({ excelApi19: false });
      const adapter = new OfficeJsAdapter();
      const got = await adapter.getSheetPageLayout("Sheet1");
      expect(got.ok).toBe(false);
      if (!got.ok) expect(got.unsupported).toBe(true);
    });
  });

  describe("WPS + prompts", () => {
    it("WPS still typed unsupported", async () => {
      const executor = new ToolExecutor(new WpsJsaAdapter());
      const result = await executor.execute({
        name: "sheet.pageLayout.set",
        arguments: { sheetName: "Sheet1", clearPageBreaks: true },
      });
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.unsupported).toBe(true);
    });

    it("boundary documents manual breaks and automatic unsupported", () => {
      const text = buildAdvancedExcelBoundary({});
      expect(text).toMatch(/manual horizontalPageBreaks\|verticalPageBreaks/);
      expect(text).toMatch(/clearPageBreaks/);
      expect(text).toMatch(/自动分页/);
      expect(text).toMatch(/margins\.\{top,bottom,left,right,header,footer\} points/);
      expect(text).toMatch(/headers\|footers default 页/);
    });
  });
});
