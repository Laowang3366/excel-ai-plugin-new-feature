import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { parsePivotDestination } from "../shared/host/officeJsPivotDestination";
import { buildPivotFieldPlan } from "../shared/host/officeJsPivotFields";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { CHAT_READONLY_TOOL_ALLOWLIST } from "../shared/agentChat/chatReadOnlyTools";
import { MockHostAdapter } from "./mockHost";
import { installPivotExcel } from "./fakes/officeJsPivotFake";

describe("phase45 pivot lifecycle", () => {
  describe("schema + registry", () => {
    it("registers pivot.list/create/refresh with closed schemas", () => {
      const names = TOOL_DEFINITIONS.map((d) => d.name);
      expect(names).toContain("pivot.list");
      expect(names).toContain("pivot.create");
      expect(names).toContain("pivot.refresh");
      expect(TOOL_DEFINITIONS).toHaveLength(83);
      for (const name of ["pivot.list", "pivot.create", "pivot.refresh"] as const) {
        const def = TOOL_DEFINITIONS.find((d) => d.name === name)!;
        expect(def.parameters.additionalProperties).toBe(false);
      }
      expect(CHAT_READONLY_TOOL_ALLOWLIST).toContain("pivot.list");
      expect(CHAT_READONLY_TOOL_ALLOWLIST).not.toContain("pivot.create");
    });

    it("rejects unknown args and missing advancedIntent at executor", async () => {
      const host = new MockHostAdapter();
      const ex = new ToolExecutor(host);
      const bad = await ex.execute({
        name: "pivot.create",
        arguments: {
          advancedIntent: "interactive-pivot",
          sourceSheetName: "Sheet1",
          sourceAddress: "A1:C10",
          extra: true,
        },
      });
      expect(bad.ok).toBe(false);
      if (!bad.ok) expect(bad.error).toMatch(/unknown field/i);

      const noIntent = await ex.execute({
        name: "pivot.create",
        arguments: { sourceSheetName: "Sheet1", sourceAddress: "A1:C10" },
      });
      expect(noIntent.ok).toBe(false);
      if (!noIntent.ok) expect(noIntent.error).toMatch(/advancedIntent/);
    });
  });

  describe("field plan + destination parse", () => {
    it("rejects axis conflicts, bad function, empty, multi-area source destination", () => {
      expect(() =>
        buildPivotFieldPlan({
          sourceSheetName: "S",
          sourceAddress: "A1",
          rowFields: ["Region"],
          columnFields: ["Region"],
        }),
      ).toThrow(/both rowFields and columnFields/);

      expect(() =>
        buildPivotFieldPlan({
          sourceSheetName: "S",
          sourceAddress: "A1",
          dataFields: [{ name: "Sales", function: "median" as "sum" }],
        }),
      ).toThrow(/sum\|count/);

      expect(() => parsePivotDestination("Sheet1!A1,B1")).toThrow(/multi-area/);
      expect(() => parsePivotDestination("[Book.xlsx]Sheet1!A1")).toThrow(/structured|external/i);
      expect(parsePivotDestination(undefined).useDedicatedSheet).toBe(true);
      expect(parsePivotDestination("A5")).toEqual({
        useDedicatedSheet: false,
        sheetName: null,
        address: "A5",
      });
      expect(parsePivotDestination("'Sheet 2'!B2").sheetName).toBe("Sheet 2");
    });
  });

  describe("MockHost + executor", () => {
    it("creates on Pivots sheet, lists, refreshes", async () => {
      const host = new MockHostAdapter();
      const ex = new ToolExecutor(host);
      const created = await ex.execute({
        name: "pivot.create",
        arguments: {
          advancedIntent: "interactive-pivot",
          sourceSheetName: "Sheet1",
          sourceAddress: "A1:C10",
          name: "SalesPivot",
          rowFields: ["Region"],
          dataFields: [{ name: "Sales", function: "sum", caption: "Sum of Sales" }],
        },
      });
      expect(created.ok).toBe(true);
      if (created.ok) {
        const data = created.data as { sheetName: string; verification: { ok: boolean } };
        expect(data.sheetName).toBe("Pivots");
        expect(data.verification.ok).toBe(true);
      }

      const listed = await ex.execute({ name: "pivot.list", arguments: {} });
      expect(listed.ok).toBe(true);
      if (listed.ok) {
        const data = listed.data as { pivots: Array<{ name: string }> };
        expect(data.pivots.map((p) => p.name)).toContain("SalesPivot");
      }

      const refreshed = await ex.execute({
        name: "pivot.refresh",
        arguments: { advancedIntent: "interactive-pivot", name: "SalesPivot" },
      });
      expect(refreshed.ok).toBe(true);
      if (refreshed.ok) {
        expect((refreshed.data as { count: number }).count).toBe(1);
      }
    });

    it("rejects refreshConnections true", async () => {
      const host = new MockHostAdapter();
      const result = await host.refreshPivots({ refreshConnections: true });
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.reason).toMatch(/refreshConnections/);
    });
  });

  describe("Office.js fake", () => {
    let fake: ReturnType<typeof installPivotExcel>;
    beforeEach(() => {
      fake = installPivotExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
    });

    it("creates with fields, lists, refreshes with readback", async () => {
      const adapter = new OfficeJsAdapter();
      const created = await adapter.createPivot({
        sourceSheetName: "Sheet1",
        sourceAddress: "A1:C10",
        name: "P1",
        rowFields: ["Region"],
        columnFields: ["Product"],
        dataFields: [{ name: "Sales", function: "sum" }],
      });
      expect(created.ok).toBe(true);
      if (created.ok) {
        expect(created.data.name).toBe("P1");
        expect(created.data.sheetName).toBe("Pivots");
        expect(created.data.verification.rowFieldCount).toBe(1);
        expect(created.data.verification.columnFieldCount).toBe(1);
        expect(created.data.verification.dataFieldCount).toBe(1);
        expect(created.data.verification.ok).toBe(true);
      }
      expect(fake.addCalls()).toBe(1);
      expect(fake.pivotNames("Pivots")).toContain("P1");

      const listed = await adapter.listPivots({});
      expect(listed.ok).toBe(true);
      if (listed.ok) {
        expect(listed.data.pivots.some((p) => p.name === "P1")).toBe(true);
      }

      const refreshed = await adapter.refreshPivots({ name: "P1" });
      expect(refreshed.ok).toBe(true);
      if (refreshed.ok) {
        expect(refreshed.data.count).toBe(1);
        expect(refreshed.data.refreshed[0]?.refreshed).toBe(true);
      }
      expect(fake.refreshCalls()).toBeGreaterThanOrEqual(1);
    });

    it("explicit destination on source sheet", async () => {
      const adapter = new OfficeJsAdapter();
      const created = await adapter.createPivot({
        sourceSheetName: "Sheet1",
        sourceAddress: "A1:C5",
        name: "OnSource",
        destination: "E1",
        rowFields: ["Region"],
        dataFields: ["Sales"],
      });
      expect(created.ok).toBe(true);
      if (created.ok) {
        expect(created.data.sheetName).toBe("Sheet1");
        expect(created.data.destination).toMatch(/E1/i);
      }
    });

    it("ExcelApi 1.8 false / missing isSetSupported → unsupported, no add", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      const f = installPivotExcel({ excelApi18: false });
      const result = await new OfficeJsAdapter().createPivot({
        sourceSheetName: "Sheet1",
        sourceAddress: "A1:C2",
        rowFields: ["Region"],
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(true);
        expect(result.reason).toMatch(/ExcelApi 1\.8/);
      }
      expect(f.addCalls()).toBe(0);

      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      const f2 = installPivotExcel({ missingIsSetSupported: true });
      const r2 = await new OfficeJsAdapter().listPivots();
      expect(r2.ok).toBe(false);
      if (!r2.ok) expect(r2.unsupported).toBe(true);
      expect(f2.addCalls()).toBe(0);
    });

    it("hierarchy tamper fails verification (no fake success)", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installPivotExcel({ tamperHierarchies: true });
      const result = await new OfficeJsAdapter().createPivot({
        sourceSheetName: "Sheet1",
        sourceAddress: "A1:C10",
        name: "Bad",
        rowFields: ["Region"],
        dataFields: ["Sales"],
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).not.toBe(true);
        expect(result.reason).toMatch(/pivot_verification_failed|rowFieldCount/i);
      }
    });

    it("rejects illegal source and refreshConnections", async () => {
      const adapter = new OfficeJsAdapter();
      const multi = await adapter.createPivot({
        sourceSheetName: "Sheet1",
        sourceAddress: "A1:B2,C3",
        rowFields: ["Region"],
      });
      expect(multi.ok).toBe(false);

      const conn = await adapter.refreshPivots({ refreshConnections: true });
      expect(conn.ok).toBe(false);
      if (!conn.ok) expect(conn.reason).toMatch(/refreshConnections/);
    });

    it("missing add member after precheck is failed not unsupported", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installPivotExcel({ missingAdd: true });
      const result = await new OfficeJsAdapter().createPivot({
        sourceSheetName: "Sheet1",
        sourceAddress: "A1:C3",
        rowFields: ["Region"],
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).not.toBe(true);
        expect(result.reason).toMatch(/pivotTables\.add|not available/i);
      }
    });
  });

  describe("WPS", () => {
    it("lists create refresh as typed unsupported", async () => {
      const wps = new WpsJsaAdapter();
      for (const result of [
        await wps.listPivots(),
        await wps.createPivot({
          sourceSheetName: "Sheet1",
          sourceAddress: "A1:C2",
        }),
        await wps.refreshPivots({}),
      ]) {
        expect(result.ok).toBe(false);
        if (!result.ok) {
          expect(result.unsupported).toBe(true);
          expect(result.reason).toMatch(/WPS|not verified/i);
        }
      }
    });
  });
});
