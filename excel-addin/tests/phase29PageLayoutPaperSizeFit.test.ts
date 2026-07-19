import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { PAGE_LAYOUT_TOOL_DEFINITIONS } from "../shared/tools/pageLayoutDefinitions";
import { ToolExecutor } from "../shared/tools";
import { installPageLayoutExcel } from "./fakes/officeJsPageLayoutFake";
import { MockHostAdapter } from "./mockHost";

const PAPER_SIZES = [
  { public: "a3" as const, host: "A3" },
  { public: "a4" as const, host: "A4" },
  { public: "a5" as const, host: "A5" },
  { public: "letter" as const, host: "Letter" },
  { public: "legal" as const, host: "Legal" },
];

describe("phase29 sheet.pageLayout paperSize + fitToPages", () => {
  describe("schema", () => {
    it("set schema exposes paperSize enum and fit integer bounds", () => {
      const setDef = PAGE_LAYOUT_TOOL_DEFINITIONS.find((t) => t.name === "sheet.pageLayout.set");
      expect(setDef).toBeDefined();
      const props = setDef!.parameters.properties as Record<string, Record<string, unknown>>;
      expect(props.paperSize?.enum).toEqual(["a3", "a4", "a5", "letter", "legal"]);
      expect(props.fitToPagesWide?.type).toBe("integer");
      expect(props.fitToPagesWide?.minimum).toBe(1);
      expect(props.fitToPagesWide?.maximum).toBe(32767);
      expect(props.fitToPagesTall?.type).toBe("integer");
      expect(props.fitToPagesTall?.minimum).toBe(1);
      expect(props.fitToPagesTall?.maximum).toBe(32767);
      expect(setDef!.parameters.additionalProperties).toBe(false);
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

    it("writes and host-readbacks all five paperSize values", async () => {
      const adapter = new OfficeJsAdapter();
      for (const { public: pub, host } of PAPER_SIZES) {
        const set = await adapter.setSheetPageLayout({ sheetName: "Sheet1", paperSize: pub });
        expect(set.ok).toBe(true);
        if (set.ok) {
          expect(set.data.paperSize).toBe(pub);
          expect(gates.getCommitted("Sheet1")?.paperSize).toBe(host);
        }
      }
    });

    it("returns unknown non-empty host paperSize as-is", async () => {
      gates.setCommittedPaperSize("Sheet1", "PaperUser");
      const adapter = new OfficeJsAdapter();
      const got = await adapter.getSheetPageLayout("Sheet1");
      expect(got.ok).toBe(true);
      if (got.ok) expect(got.data.paperSize).toBe("PaperUser");
    });

    it("returns host Worksheet.name, not input sheetName echo", async () => {
      gates.setHostSheetName("Sheet1", "HostSheetAlpha");
      const adapter = new OfficeJsAdapter();
      const got = await adapter.getSheetPageLayout("Sheet1");
      expect(got.ok).toBe(true);
      if (got.ok) {
        expect(got.data.sheetName).toBe("HostSheetAlpha");
        expect(got.data.sheetName).not.toBe("Sheet1");
      }
    });

    it("fails on empty or non-string host paperSize as ordinary failure", async () => {
      gates.setCommittedPaperSize("Sheet1", "");
      const adapter = new OfficeJsAdapter();
      const empty = await adapter.getSheetPageLayout("Sheet1");
      expect(empty.ok).toBe(false);
      if (!empty.ok) {
        expect(empty.reason).toMatch(/paperSize|non-empty string/i);
        expect(empty.reason).not.toMatch(/is not supported in this host|isSetSupported/i);
        expect(empty.unsupported).not.toBe(true);
      }
      expect(gates.getExcelRunCalls()).toBe(1);

      gates.setCommittedPaperSize("Sheet1", null as unknown as string);
      const bad = await adapter.getSheetPageLayout("Sheet1");
      expect(bad.ok).toBe(false);
      if (!bad.ok) {
        expect(bad.reason).toMatch(/paperSize|non-empty string/i);
        expect(bad.unsupported).not.toBe(true);
      }
    });

    it("fit boundaries 1 and 32767, wide-only, tall-only, both", async () => {
      const adapter = new OfficeJsAdapter();

      const wideOnly = await adapter.setSheetPageLayout({
        sheetName: "Sheet1",
        fitToPagesWide: 1,
      });
      expect(wideOnly.ok).toBe(true);
      if (wideOnly.ok) {
        expect(wideOnly.data.fitToPagesWide).toBe(1);
        expect(wideOnly.data.fitToPagesTall).toBeNull();
        expect(wideOnly.data.zoomScale).toBeNull();
      }

      const tallOnly = await adapter.setSheetPageLayout({
        sheetName: "Sheet1",
        fitToPagesTall: 32767,
      });
      expect(tallOnly.ok).toBe(true);
      if (tallOnly.ok) {
        expect(tallOnly.data.fitToPagesTall).toBe(32767);
        expect(tallOnly.data.zoomScale).toBeNull();
      }

      const both = await adapter.setSheetPageLayout({
        sheetName: "Sheet1",
        fitToPagesWide: 2,
        fitToPagesTall: 3,
      });
      expect(both.ok).toBe(true);
      if (both.ok) {
        expect(both.data.fitToPagesWide).toBe(2);
        expect(both.data.fitToPagesTall).toBe(3);
        expect(both.data.zoomScale).toBeNull();
        expect(gates.getCommitted("Sheet1")?.fitToPagesWide).toBe(2);
        expect(gates.getCommitted("Sheet1")?.fitToPagesTall).toBe(3);
        expect(gates.getCommitted("Sheet1")?.zoomScale).toBeNull();
      }
    });

    it("zoomScale write clears fit fields; fit write clears scale", async () => {
      const adapter = new OfficeJsAdapter();
      const fit = await adapter.setSheetPageLayout({
        sheetName: "Sheet1",
        fitToPagesWide: 1,
        fitToPagesTall: 1,
      });
      expect(fit.ok).toBe(true);
      if (fit.ok) expect(fit.data.zoomScale).toBeNull();

      const scale = await adapter.setSheetPageLayout({
        sheetName: "Sheet1",
        zoomScale: 150,
      });
      expect(scale.ok).toBe(true);
      if (scale.ok) {
        expect(scale.data.zoomScale).toBe(150);
        expect(scale.data.fitToPagesWide).toBeNull();
        expect(scale.data.fitToPagesTall).toBeNull();
      }

      const fitAgain = await adapter.setSheetPageLayout({
        sheetName: "Sheet1",
        fitToPagesWide: 4,
      });
      expect(fitAgain.ok).toBe(true);
      if (fitAgain.ok) {
        expect(fitAgain.data.fitToPagesWide).toBe(4);
        expect(fitAgain.data.zoomScale).toBeNull();
      }
    });

    it("sync gate: paperSize/zoom stay stale until sync (no input echo)", async () => {
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
        const sheet = context.workbook.worksheets.getItem("Sheet1") as {
          pageLayout: {
            paperSize: string;
            zoom: { scale?: number | null };
          };
        };
        sheet.pageLayout.paperSize = "A3";
        sheet.pageLayout.zoom = { scale: 200 };
        expect(gates.getCommitted("Sheet1")?.paperSize).toBe("Letter");
        expect(gates.getCommitted("Sheet1")?.zoomScale).toBe(100);
        expect(gates.getPending("Sheet1")?.paperSize).toBe("A3");
        expect(gates.getPending("Sheet1")?.zoomScale).toBe(200);
        await context.sync();
        expect(gates.getCommitted("Sheet1")?.paperSize).toBe("A3");
        expect(gates.getCommitted("Sheet1")?.zoomScale).toBe(200);
        expect(gates.getPending("Sheet1")).toBeUndefined();
      });
    });

    it("null zoom.scale host readback stays null (not coerced to 0)", async () => {
      gates.setCommittedZoomScale("Sheet1", null);
      const adapter = new OfficeJsAdapter();
      const got = await adapter.getSheetPageLayout("Sheet1");
      expect(got.ok).toBe(true);
      if (got.ok) {
        expect(got.data.zoomScale).toBeNull();
        expect(got.data.zoomScale).not.toBe(0);
      }
    });

    it("fit host readback when scale is null", async () => {
      gates.setCommittedFit("Sheet1", 2, 3);
      const adapter = new OfficeJsAdapter();
      const got = await adapter.getSheetPageLayout("Sheet1");
      expect(got.ok).toBe(true);
      if (got.ok) {
        expect(got.data.zoomScale).toBeNull();
        expect(got.data.fitToPagesWide).toBe(2);
        expect(got.data.fitToPagesTall).toBe(3);
      }
    });
  });

  describe("executor validation", () => {
    it("rejects fit/zoom mutual exclusion and invalid fit/paperSize", async () => {
      const executor = new ToolExecutor(new MockHostAdapter());

      const okWide = await executor.execute({
        name: "sheet.pageLayout.set",
        arguments: { sheetName: "Sheet1", fitToPagesWide: 1 },
      });
      expect(okWide.ok).toBe(true);

      const okPaper = await executor.execute({
        name: "sheet.pageLayout.set",
        arguments: { sheetName: "Sheet1", paperSize: "a4" },
      });
      expect(okPaper.ok).toBe(true);

      for (const args of [
        { sheetName: "Sheet1", zoomScale: 100, fitToPagesWide: 1 },
        { sheetName: "Sheet1", zoomScale: 100, fitToPagesTall: 1 },
        { sheetName: "Sheet1", fitToPagesWide: 0 },
        { sheetName: "Sheet1", fitToPagesWide: 32768 },
        { sheetName: "Sheet1", fitToPagesTall: 1.5 },
        { sheetName: "Sheet1", fitToPagesWide: Number.NaN },
        { sheetName: "Sheet1", fitToPagesWide: Number.POSITIVE_INFINITY },
        { sheetName: "Sheet1", fitToPagesTall: Number.NEGATIVE_INFINITY },
        { sheetName: "Sheet1", fitToPagesWide: null },
        { sheetName: "Sheet1", fitToPagesTall: undefined },
        { sheetName: "Sheet1", fitToPagesWide: "1" },
        { sheetName: "Sheet1", paperSize: "A4" },
        { sheetName: "Sheet1", paperSize: "tabloid" },
        { sheetName: "Sheet1", paperSize: null },
        { sheetName: "Sheet1", paperSize: undefined },
        { sheetName: "Sheet1", unknownKey: true },
      ]) {
        const result = await executor.execute({
          name: "sheet.pageLayout.set",
          arguments: args as Record<string, unknown>,
        });
        expect(result.ok).toBe(false);
      }
    });

    it("MockHost parity for paperSize and fit fields", async () => {
      const executor = new ToolExecutor(new MockHostAdapter());
      const set = await executor.execute({
        name: "sheet.pageLayout.set",
        arguments: {
          sheetName: "Sheet1",
          paperSize: "legal",
          fitToPagesWide: 2,
          fitToPagesTall: 3,
        },
      });
      expect(set.ok).toBe(true);
      if (set.ok) {
        const data = set.data as {
          paperSize: string;
          fitToPagesWide: number | null;
          fitToPagesTall: number | null;
          zoomScale: number | null;
        };
        expect(data.paperSize).toBe("legal");
        expect(data.fitToPagesWide).toBe(2);
        expect(data.fitToPagesTall).toBe(3);
        expect(data.zoomScale).toBeNull();
      }
    });
  });
});
