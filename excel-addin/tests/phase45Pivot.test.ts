import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { MockHostAdapter } from "./mockHost";
import { installPivotExcel } from "./fakes/officeJsPivotFake";

describe("phase45 pivot host paths", () => {
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

    it("rejects zero-field create", async () => {
      const host = new MockHostAdapter();
      const empty = await host.createPivot({
        sourceSheetName: "Sheet1",
        sourceAddress: "A1:C2",
      });
      expect(empty.ok).toBe(false);
      if (!empty.ok) expect(empty.reason).toMatch(/at least one field/);
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
      delete (globalThis as { window?: unknown }).window;
    });

    it("creates with multi data agg, lists, refreshes with readback", async () => {
      const adapter = new OfficeJsAdapter();
      const created = await adapter.createPivot({
        sourceSheetName: "Sheet1",
        sourceAddress: "A1:C10",
        name: "P1",
        rowFields: ["Region"],
        columnFields: ["Product"],
        dataFields: [
          { name: "Sales", function: "sum", caption: "Sum Sales" },
          { name: "Sales", function: "count", caption: "Count Sales" },
        ],
      });
      expect(created.ok).toBe(true);
      if (created.ok) {
        expect(created.data.name).toBe("P1");
        expect(created.data.sheetName).toBe("Pivots");
        expect(created.data.verification.rowFieldCount).toBe(1);
        expect(created.data.verification.columnFieldCount).toBe(1);
        expect(created.data.verification.dataFieldCount).toBe(2);
        expect(created.data.verification.ok).toBe(true);
      }
      expect(fake.addCalls()).toBe(1);

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

    it("explicit destination; zero-field fails closed before add", async () => {
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

      const empty = await adapter.createPivot({
        sourceSheetName: "Sheet1",
        sourceAddress: "A1:C5",
        name: "Empty",
      });
      expect(empty.ok).toBe(false);
      if (!empty.ok) {
        expect(empty.unsupported).not.toBe(true);
        expect(empty.reason).toMatch(/at least one field/);
      }
      expect(fake.addCalls()).toBe(1);
    });

    it("ExcelApi 1.8 false blocks create; 1.3 false blocks refresh", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      const f = installPivotExcel({ excelApi18: false });
      const create = await new OfficeJsAdapter().createPivot({
        sourceSheetName: "Sheet1",
        sourceAddress: "A1:C2",
        rowFields: ["Region"],
      });
      expect(create.ok).toBe(false);
      if (!create.ok) {
        expect(create.unsupported).toBe(true);
        expect(create.reason).toMatch(/ExcelApi 1\.8/);
      }
      expect(f.addCalls()).toBe(0);

      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      const f13 = installPivotExcel({ excelApi18: true, excelApi13: false });
      f13.seedPivot({ name: "R1", sheetName: "Sheet1" });
      const refresh = await new OfficeJsAdapter().refreshPivots({ name: "R1" });
      expect(refresh.ok).toBe(false);
      if (!refresh.ok) {
        expect(refresh.unsupported).toBe(true);
        expect(refresh.reason).toMatch(/ExcelApi 1\.3/);
      }
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
        expect(result.reason).toMatch(/pivot_verification_failed|rowFieldCount|hasFields/i);
      }
    });

    it("rejects illegal multi-area source", async () => {
      const adapter = new OfficeJsAdapter();
      const multi = await adapter.createPivot({
        sourceSheetName: "Sheet1",
        sourceAddress: "A1:B2,C3",
        rowFields: ["Region"],
      });
      expect(multi.ok).toBe(false);
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
          rowFields: ["Region"],
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

// Keep registry size assertion reachable from host suite as well.
describe("phase45 registry smoke", () => {
  it("keeps 89 tools after markers", () => {
    expect(TOOL_DEFINITIONS).toHaveLength(89);
  });
});
