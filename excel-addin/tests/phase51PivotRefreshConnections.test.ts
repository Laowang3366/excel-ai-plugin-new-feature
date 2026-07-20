import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { ToolExecutor } from "../shared/tools/executor";
import { TOOL_DEFINITIONS } from "../shared/tools/definitions";
import { installPivotExcel } from "./fakes/officeJsPivotFake";
import { MockHostAdapter } from "./mockHost";

describe("phase51 pivot.refresh refreshConnections", () => {
  describe("Office.js", () => {
    let fake: ReturnType<typeof installPivotExcel>;
    beforeEach(() => {
      fake = installPivotExcel();
      fake.seedPivot({ name: "P1", sheetName: "Sheet1", row: ["Region"] });
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
    });

    it("true refreshes pivots and queues dataConnections.refreshAll with verified:false", async () => {
      const result = await new OfficeJsAdapter().refreshPivots({
        refreshConnections: true,
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.count).toBe(1);
        expect(result.data.refreshed[0]?.name).toBe("P1");
        expect(result.data.connectionRefresh).toEqual({
          requested: true,
          method: "Workbook.dataConnections.refreshAll",
          verified: false,
          scope: "supported-office-js-connections",
        });
        expect(result.data.limitations?.join(" ")).toMatch(/Power Query/i);
        expect(result.data.limitations?.join(" ")).toMatch(/firewall/i);
        expect(result.data.limitations?.join(" ")).toMatch(/verified:false|readback/i);
      }
      expect(fake.refreshCalls()).toBeGreaterThanOrEqual(1);
      expect(fake.connectionRefreshCalls()).toBe(1);
    });

    it("false/omitted never calls dataConnections.refreshAll", async () => {
      const a = await new OfficeJsAdapter().refreshPivots({});
      expect(a.ok).toBe(true);
      expect(fake.connectionRefreshCalls()).toBe(0);
      const b = await new OfficeJsAdapter().refreshPivots({ refreshConnections: false });
      expect(b.ok).toBe(true);
      if (b.ok) expect(b.data.connectionRefresh).toBeUndefined();
      expect(fake.connectionRefreshCalls()).toBe(0);
      expect(fake.refreshCalls()).toBeGreaterThanOrEqual(2);
    });

    it("ExcelApi 1.7 false/missing/throw → typed unsupported and zero refresh side effects", async () => {
      for (const opts of [
        { excelApi17: false },
        { missingIsSetSupported: true },
        { isSetSupportedThrows: true },
      ] as const) {
        delete (globalThis as { Excel?: unknown }).Excel;
        delete (globalThis as { Office?: unknown }).Office;
        fake = installPivotExcel(opts);
        fake.seedPivot({ name: "P1", sheetName: "Sheet1" });
        const result = await new OfficeJsAdapter().refreshPivots({ refreshConnections: true });
        expect(result.ok).toBe(false);
        if (!result.ok) {
          expect(result.unsupported).toBe(true);
          expect(result.reason).toMatch(/ExcelApi 1\.7|isSetSupported/i);
        }
        expect(fake.refreshCalls()).toBe(0);
        expect(fake.connectionRefreshCalls()).toBe(0);
      }
    });

    it("1.7 true but missing dataConnections/refreshAll is ordinary failed", async () => {
      for (const opts of [{ missingDataConnections: true }, { missingRefreshAll: true }] as const) {
        delete (globalThis as { Excel?: unknown }).Excel;
        delete (globalThis as { Office?: unknown }).Office;
        fake = installPivotExcel(opts);
        fake.seedPivot({ name: "P1", sheetName: "Sheet1" });
        const result = await new OfficeJsAdapter().refreshPivots({ refreshConnections: true });
        expect(result.ok).toBe(false);
        if (!result.ok) {
          expect(result.unsupported).not.toBe(true);
          expect(result.reason).toMatch(/dataConnections\.refreshAll/i);
        }
        expect(fake.connectionRefreshCalls()).toBe(0);
        expect(fake.refreshCalls()).toBe(0);
      }
    });

    it("explicit missing pivot fails without connection refresh", async () => {
      const result = await new OfficeJsAdapter().refreshPivots({
        name: "Missing",
        refreshConnections: true,
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).not.toBe(true);
        expect(result.reason).toMatch(/not found/i);
      }
      expect(fake.refreshCalls()).toBe(0);
      expect(fake.connectionRefreshCalls()).toBe(0);
    });

    it("zero pivots + true only requests connection refresh", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      fake = installPivotExcel();
      const result = await new OfficeJsAdapter().refreshPivots({ refreshConnections: true });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.count).toBe(0);
        expect(result.data.refreshed).toEqual([]);
        expect(result.data.connectionRefresh?.requested).toBe(true);
        expect(result.data.connectionRefresh?.verified).toBe(false);
      }
      expect(fake.refreshCalls()).toBe(0);
      expect(fake.connectionRefreshCalls()).toBe(1);
    });

    it("zero pivots + false keeps no-pivot limitation", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      fake = installPivotExcel();
      const result = await new OfficeJsAdapter().refreshPivots({});
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.count).toBe(0);
        expect(result.data.limitations?.join(" ")).toMatch(/no pivot/i);
        expect(result.data.connectionRefresh).toBeUndefined();
      }
      expect(fake.connectionRefreshCalls()).toBe(0);
    });
  });

  describe("executor / schema / WPS", () => {
    it("rejects null/unknown fields before Host", async () => {
      const host = new MockHostAdapter();
      let calls = 0;
      const original = host.refreshPivots.bind(host);
      host.refreshPivots = async (input) => {
        calls += 1;
        return original(input);
      };
      const ex = new ToolExecutor(host);
      for (const args of [
        { advancedIntent: "interactive-pivot", refreshConnections: null },
        { advancedIntent: "interactive-pivot", extra: true },
        { advancedIntent: "interactive-pivot", sheetName: null },
      ]) {
        const result = await ex.execute({
          name: "pivot.refresh",
          arguments: args as Record<string, unknown>,
        });
        expect(result.ok).toBe(false);
      }
      expect(calls).toBe(0);
    });

    it("mock host accepts refreshConnections true with verified false", async () => {
      const host = new MockHostAdapter();
      await host.createPivot({
        sourceSheetName: "Sheet1",
        sourceAddress: "A1:C3",
        rowFields: ["Region"],
        name: "M1",
      });
      const result = await new ToolExecutor(host).execute({
        name: "pivot.refresh",
        arguments: {
          advancedIntent: "interactive-pivot",
          refreshConnections: true,
        },
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        const data = result.data as {
          connectionRefresh?: { verified: boolean; method: string };
          count: number;
        };
        expect(data.count).toBe(1);
        expect(data.connectionRefresh?.verified).toBe(false);
        expect(data.connectionRefresh?.method).toBe("Workbook.dataConnections.refreshAll");
      }
    });

    it("WPS remains typed unsupported; tool count stays 89", async () => {
      const result = await new ToolExecutor(new WpsJsaAdapter()).execute({
        name: "pivot.refresh",
        arguments: {
          advancedIntent: "interactive-pivot",
          refreshConnections: true,
        },
      });
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.unsupported).toBe(true);
      expect(TOOL_DEFINITIONS).toHaveLength(89);
    });
  });
});
