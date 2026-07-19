import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { buildAdvancedExcelBoundary } from "../shared/prompts/advancedExcelBoundary";
import { PAGE_LAYOUT_TOOL_DEFINITIONS } from "../shared/tools/pageLayoutDefinitions";
import { ToolExecutor } from "../shared/tools";
import { installPageLayoutExcel } from "./fakes/officeJsPageLayoutFake";
import { MockHostAdapter } from "./mockHost";

const SIDE_SCHEMA = { type: "string" };

describe("phase32 sheet.pageLayout headers/footers default text", () => {
  describe("schema", () => {
    it("exposes headers/footers three sides with additionalProperties false", () => {
      const setDef = PAGE_LAYOUT_TOOL_DEFINITIONS.find((t) => t.name === "sheet.pageLayout.set");
      expect(setDef).toBeDefined();
      const props = setDef!.parameters.properties as Record<string, Record<string, unknown>>;
      for (const key of ["headers", "footers"] as const) {
        const side = props[key] as {
          properties?: Record<string, unknown>;
          additionalProperties?: boolean;
        };
        expect(side.properties?.left).toEqual(SIDE_SCHEMA);
        expect(side.properties?.center).toEqual(SIDE_SCHEMA);
        expect(side.properties?.right).toEqual(SIDE_SCHEMA);
        expect(side.additionalProperties).toBe(false);
      }
    });
  });

  describe("executor + MockHost", () => {
    it("accepts partial update and empty-string clear; preserves unwritten sides", async () => {
      const executor = new ToolExecutor(new MockHostAdapter());
      const seed = await executor.execute({
        name: "sheet.pageLayout.set",
        arguments: {
          sheetName: "Sheet1",
          headers: { left: "L", center: "C", right: "R" },
          footers: { left: "fl", center: "fc", right: "fr" },
        },
      });
      expect(seed.ok).toBe(true);

      const partial = await executor.execute({
        name: "sheet.pageLayout.set",
        arguments: { sheetName: "Sheet1", headers: { center: "C2" }, footers: { right: "" } },
      });
      expect(partial.ok).toBe(true);
      if (partial.ok) {
        const data = partial.data as {
          headers: { left: string; center: string; right: string };
          footers: { left: string; center: string; right: string };
        };
        expect(data.headers).toEqual({ left: "L", center: "C2", right: "R" });
        expect(data.footers).toEqual({ left: "fl", center: "fc", right: "" });
      }
    });

    it("rejects invalid headers/footers shapes and top-level header", async () => {
      const executor = new ToolExecutor(new MockHostAdapter());
      const cases: Array<Record<string, unknown>> = [
        { sheetName: "Sheet1", headers: null },
        { sheetName: "Sheet1", footers: ["x"] },
        { sheetName: "Sheet1", headers: { left: 1 } },
        { sheetName: "Sheet1", headers: { middle: "x" } },
        { sheetName: "Sheet1", headers: {} },
        { sheetName: "Sheet1", headers: { left: undefined } },
        { sheetName: "Sheet1", header: "legacy" },
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

    it("reads non-default six slots and partial set with committed host state", async () => {
      gates.setCommittedHeadersFooters("Sheet1", {
        leftHeader: "HL",
        centerHeader: "HC",
        rightHeader: "HR",
        leftFooter: "FL",
        centerFooter: "FC",
        rightFooter: "FR",
      });
      const adapter = new OfficeJsAdapter();
      const got = await adapter.getSheetPageLayout("Sheet1");
      expect(got.ok).toBe(true);
      if (got.ok) {
        expect(got.data.headers).toEqual({ left: "HL", center: "HC", right: "HR" });
        expect(got.data.footers).toEqual({ left: "FL", center: "FC", right: "FR" });
      }

      const set = await adapter.setSheetPageLayout({
        sheetName: "Sheet1",
        headers: { left: "L2" },
        footers: { center: "FC2" },
      });
      expect(set.ok).toBe(true);
      if (set.ok) {
        expect(set.data.headers).toEqual({ left: "L2", center: "HC", right: "HR" });
        expect(set.data.footers).toEqual({ left: "FL", center: "FC2", right: "FR" });
      }
      const committed = gates.getCommitted("Sheet1");
      expect(committed?.leftHeader).toBe("L2");
      expect(committed?.centerHeader).toBe("HC");
      expect(committed?.centerFooter).toBe("FC2");

      // prove not input-echo: rename host sheet name
      gates.setHostSheetName("Sheet1", "HostHF");
      const again = await adapter.getSheetPageLayout("Sheet1");
      expect(again.ok).toBe(true);
      if (again.ok) {
        expect(again.data.sheetName).toBe("HostHF");
        expect(again.data.headers.left).toBe("L2");
      }
    });

    it("clears a side with empty string after non-empty write", async () => {
      const adapter = new OfficeJsAdapter();
      const filled = await adapter.setSheetPageLayout({
        sheetName: "Sheet1",
        headers: { right: "X" },
      });
      expect(filled.ok).toBe(true);
      expect(gates.getCommitted("Sheet1")?.rightHeader).toBe("X");

      const cleared = await adapter.setSheetPageLayout({
        sheetName: "Sheet1",
        headers: { right: "" },
      });
      expect(cleared.ok).toBe(true);
      if (cleared.ok) expect(cleared.data.headers.right).toBe("");
      expect(gates.getCommitted("Sheet1")?.rightHeader).toBe("");
    });

    it("missing headersFooters/defaultForAllPages/slot is ordinary failure", async () => {
      for (const opts of [
        { hasHeadersFooters: false },
        { hasDefaultForAllPages: false },
        { missingHeaderFooterSlot: "leftHeader" as const },
        { missingHeaderFooterSlot: "centerFooter" as const },
      ]) {
        delete (globalThis as { Excel?: unknown }).Excel;
        delete (globalThis as { Office?: unknown }).Office;
        installPageLayoutExcel(opts);
        const adapter = new OfficeJsAdapter();
        const got = await adapter.getSheetPageLayout("Sheet1");
        expect(got.ok).toBe(false);
        if (!got.ok) {
          expect(got.unsupported).not.toBe(true);
          expect(got.capability).toBe("sheet.pageLayout.get");
          expect(got.host).toBe("office-js");
        }

        const set = await adapter.setSheetPageLayout({
          sheetName: "Sheet1",
          headers: { left: "z" },
        });
        expect(set.ok).toBe(false);
        if (!set.ok) {
          expect(set.unsupported).not.toBe(true);
          expect(set.capability).toBe("sheet.pageLayout.set");
          expect(set.host).toBe("office-js");
        }
      }
    });

    it("host slot null/number is ordinary failure", async () => {
      const adapter = new OfficeJsAdapter();
      gates.setCommittedHeadersFooters("Sheet1", { leftHeader: null as unknown as string });
      const nullSlot = await adapter.getSheetPageLayout("Sheet1");
      expect(nullSlot.ok).toBe(false);
      if (!nullSlot.ok) {
        expect(nullSlot.unsupported).not.toBe(true);
        expect(nullSlot.reason).toMatch(/leftHeader|string/i);
      }

      gates.setCommittedHeadersFooters("Sheet1", {
        leftHeader: "",
        centerHeader: 42 as unknown as string,
      });
      const numSlot = await adapter.getSheetPageLayout("Sheet1");
      expect(numSlot.ok).toBe(false);
      if (!numSlot.ok) {
        expect(numSlot.unsupported).not.toBe(true);
        expect(numSlot.reason).toMatch(/centerHeader|string/i);
      }
    });

    it("ExcelApi 1.9 off remains typed unsupported", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installPageLayoutExcel({ excelApi19: false });
      const adapter = new OfficeJsAdapter();
      const got = await adapter.getSheetPageLayout("Sheet1");
      expect(got.ok).toBe(false);
      if (!got.ok) {
        expect(got.unsupported).toBe(true);
        expect(got.capability).toBe("sheet.pageLayout.get");
      }
    });
  });

  describe("WPS + prompts", () => {
    it("WPS still typed unsupported", async () => {
      const executor = new ToolExecutor(new WpsJsaAdapter());
      for (const name of ["sheet.pageLayout.get", "sheet.pageLayout.set"] as const) {
        const result = await executor.execute({
          name,
          arguments:
            name === "sheet.pageLayout.get"
              ? { sheetName: "Sheet1" }
              : { sheetName: "Sheet1", headers: { left: "x" } },
        });
        expect(result.ok).toBe(false);
        if (!result.ok) expect(result.unsupported).toBe(true);
      }
    });

    it("boundary documents default six-slot text support and first/even/odd unsupported", () => {
      const text = buildAdvancedExcelBoundary({});
      expect(text).toMatch(/headers\|footers default 页 left\|center\|right 文本/);
      expect(text).toMatch(/first\/even\/odd pages/);
      expect(text).toMatch(/margins\.\{top,bottom,left,right,header,footer\} points/);
      expect(text).not.toMatch(/headers\/footers 文本内容\/page breaks/);
    });
  });
});
