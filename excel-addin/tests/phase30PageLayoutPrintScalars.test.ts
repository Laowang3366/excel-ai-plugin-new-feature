import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { PAGE_LAYOUT_TOOL_DEFINITIONS } from "../shared/tools/pageLayoutDefinitions";
import { ToolExecutor } from "../shared/tools";
import { installPageLayoutExcel } from "./fakes/officeJsPageLayoutFake";
import { MockHostAdapter } from "./mockHost";

describe("phase30 sheet.pageLayout draft/pageOrder/firstPageNumber", () => {
  describe("schema", () => {
    it("set schema exposes pageOrder enum and firstPageNumber minimum", () => {
      const setDef = PAGE_LAYOUT_TOOL_DEFINITIONS.find((t) => t.name === "sheet.pageLayout.set");
      expect(setDef).toBeDefined();
      const props = setDef!.parameters.properties as Record<string, Record<string, unknown>>;
      expect(props.draft?.type).toBe("boolean");
      expect(props.pageOrder?.enum).toEqual(["downThenOver", "overThenDown"]);
      expect(props.firstPageNumber?.type).toBe("integer");
      expect(props.firstPageNumber?.minimum).toBe(1);
      expect(setDef!.parameters.additionalProperties).toBe(false);
    });
  });

  describe("executor validation", () => {
    it("accepts draft:false, both pageOrders, and firstPageNumber=1", async () => {
      const executor = new ToolExecutor(new MockHostAdapter());
      for (const args of [
        { sheetName: "Sheet1", draft: false },
        { sheetName: "Sheet1", pageOrder: "downThenOver" },
        { sheetName: "Sheet1", pageOrder: "overThenDown" },
        { sheetName: "Sheet1", firstPageNumber: 1 },
      ]) {
        const result = await executor.execute({ name: "sheet.pageLayout.set", arguments: args });
        expect(result.ok).toBe(true);
      }
    });

    it("rejects unknown pageOrder, bad firstPageNumber, and unknown fields", async () => {
      const executor = new ToolExecutor(new MockHostAdapter());
      const cases: Array<{ args: Record<string, unknown>; match: RegExp }> = [
        { args: { sheetName: "Sheet1", pageOrder: "sideways" }, match: /pageOrder/i },
        { args: { sheetName: "Sheet1", firstPageNumber: 0 }, match: /firstPageNumber/i },
        { args: { sheetName: "Sheet1", firstPageNumber: -1 }, match: /firstPageNumber/i },
        { args: { sheetName: "Sheet1", firstPageNumber: 1.5 }, match: /firstPageNumber/i },
        { args: { sheetName: "Sheet1", firstPageNumber: Number.NaN }, match: /firstPageNumber/i },
        { args: { sheetName: "Sheet1", firstPageNumber: Number.POSITIVE_INFINITY }, match: /firstPageNumber/i },
        { args: { sheetName: "Sheet1", firstPageNumber: null }, match: /firstPageNumber/i },
        { args: { sheetName: "Sheet1", firstPageNumber: "1" }, match: /firstPageNumber/i },
        { args: { sheetName: "Sheet1", draft: "false" }, match: /boolean|draft/i },
        { args: { sheetName: "Sheet1", draftMode: true }, match: /unknown field/i },
        { args: { sheetName: "Sheet1", headers: "x" }, match: /unknown field/i },
      ];
      for (const { args, match } of cases) {
        const result = await executor.execute({ name: "sheet.pageLayout.set", arguments: args });
        expect(result.ok).toBe(false);
        if (!result.ok) expect(result.error).toMatch(match);
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

    it("sets draft/pageOrder/firstPageNumber with real fake host readback (not input echo)", async () => {
      const adapter = new OfficeJsAdapter();
      const set = await adapter.setSheetPageLayout({
        sheetName: "Sheet1",
        draft: false,
        pageOrder: "overThenDown",
        firstPageNumber: 3,
      });
      expect(set.ok).toBe(true);
      if (set.ok) {
        expect(set.data.draft).toBe(false);
        expect(set.data.pageOrder).toBe("overThenDown");
        expect(set.data.firstPageNumber).toBe(3);
        expect(set.data.sheetName).toBe("Sheet1");
      }
      const committed = gates.getCommitted("Sheet1");
      expect(committed?.draftMode).toBe(false);
      expect(committed?.printOrder).toBe("OverThenDown");
      expect(committed?.firstPageNumber).toBe(3);

      gates.setHostSheetName("Sheet1", "HostRenamed");
      const got = await adapter.getSheetPageLayout("Sheet1");
      expect(got.ok).toBe(true);
      if (got.ok) {
        expect(got.data.sheetName).toBe("HostRenamed");
        expect(got.data.draft).toBe(false);
        expect(got.data.pageOrder).toBe("overThenDown");
        expect(got.data.firstPageNumber).toBe(3);
      }
    });

    it("maps both pageOrder values and firstPageNumber empty/null to null", async () => {
      const adapter = new OfficeJsAdapter();
      const down = await adapter.setSheetPageLayout({
        sheetName: "Sheet1",
        pageOrder: "downThenOver",
      });
      expect(down.ok).toBe(true);
      if (down.ok) expect(down.data.pageOrder).toBe("downThenOver");
      expect(gates.getCommitted("Sheet1")?.printOrder).toBe("DownThenOver");

      const over = await adapter.setSheetPageLayout({
        sheetName: "Sheet1",
        pageOrder: "overThenDown",
      });
      expect(over.ok).toBe(true);
      if (over.ok) expect(over.data.pageOrder).toBe("overThenDown");
      expect(gates.getCommitted("Sheet1")?.printOrder).toBe("OverThenDown");

      gates.setCommittedFirstPageNumber("Sheet1", "");
      const empty = await adapter.getSheetPageLayout("Sheet1");
      expect(empty.ok).toBe(true);
      if (empty.ok) expect(empty.data.firstPageNumber).toBeNull();

      gates.setCommittedFirstPageNumber("Sheet1", null);
      const nulled = await adapter.getSheetPageLayout("Sheet1");
      expect(nulled.ok).toBe(true);
      if (nulled.ok) expect(nulled.data.firstPageNumber).toBeNull();

      gates.setCommittedFirstPageNumber("Sheet1", 7);
      const numbered = await adapter.getSheetPageLayout("Sheet1");
      expect(numbered.ok).toBe(true);
      if (numbered.ok) expect(numbered.data.firstPageNumber).toBe(7);
    });

    it("missing draftMode/printOrder/firstPageNumber after 1.9 is ordinary failure", async () => {
      for (const opts of [
        { hasDraftMode: false },
        { hasPrintOrder: false },
        { hasFirstPageNumber: false },
      ] as const) {
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
          expect(got.reason).toMatch(/missing on host layout object/i);
        }
        expect(local.getExcelRunCalls()).toBe(1);
      }
    });

    it("unknown host printOrder is ordinary failure", async () => {
      gates.setCommittedPrintOrder("Sheet1", "ZigZag");
      const adapter = new OfficeJsAdapter();
      const got = await adapter.getSheetPageLayout("Sheet1");
      expect(got.ok).toBe(false);
      if (!got.ok) {
        expect(got.unsupported).not.toBe(true);
        expect(got.capability).toBe("sheet.pageLayout.get");
        expect(got.host).toBe("office-js");
        expect(got.reason).toMatch(/printOrder|unknown host value/i);
      }
    });

    it("host firstPageNumber 0/fraction/NaN/string fails ordinarily", async () => {
      const adapter = new OfficeJsAdapter();
      for (const value of [0, 1.5, Number.NaN, Number.POSITIVE_INFINITY, "auto", true]) {
        gates.setCommittedFirstPageNumber("Sheet1", value as never);
        const got = await adapter.getSheetPageLayout("Sheet1");
        expect(got.ok).toBe(false);
        if (!got.ok) {
          expect(got.unsupported).not.toBe(true);
          expect(got.capability).toBe("sheet.pageLayout.get");
          expect(got.host).toBe("office-js");
          expect(got.reason).toMatch(/firstPageNumber/i);
        }
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
              : { sheetName: "Sheet1", draft: true },
        });
        expect(result.ok).toBe(false);
        if (!result.ok) {
          expect(result.unsupported).toBe(true);
        }
      }
    });
  });
});
