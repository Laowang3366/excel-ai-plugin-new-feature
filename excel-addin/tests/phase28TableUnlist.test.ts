import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { installTableUnlistExcel } from "./fakes/officeJsTableUnlistFake";
import { MockHostAdapter } from "./mockHost";

describe("phase28 table.unlist", () => {
  describe("Office.js sync-gated", () => {
    let fake: ReturnType<typeof installTableUnlistExcel>;
    beforeEach(() => {
      fake = installTableUnlistExcel({
        hostSheetName: "HostSheet",
        tableName: "Sales",
        hostTableName: "HostSales",
        address: "HostSheet!A1:B3",
      });
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
    });

    it("returns host sheet/table/address + unlisted after absence check", async () => {
      const result = await new OfficeJsAdapter().unlistTable("Sheet1", "Sales");
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data).toEqual({
          sheetName: "HostSheet",
          tableName: "HostSales",
          address: "HostSheet!A1:B3",
          unlisted: true,
        });
        expect(result.data.sheetName).not.toBe("Sheet1");
        expect(result.data.tableName).not.toBe("Sales");
      }
      expect(fake.convertCalls()).toBe(1);
      expect(fake.tableNames()).not.toContain("HostSales");
      expect(fake.cellValuesAt("A1")).toEqual([
        ["H1", "H2"],
        ["a", 1],
        ["b", 2],
      ]);
    });

    it("skip convert sync still exposes old table (no echo success)", async () => {
      const names = await fake.brokenSkipConvertSync();
      expect(names.map((n) => n.toLowerCase())).toContain("hostsales");
      expect(fake.convertCalls()).toBe(1);
      const ok = await new OfficeJsAdapter().unlistTable("Sheet1", "Sales");
      expect(ok.ok).toBe(true);
      if (ok.ok) expect(ok.data.unlisted).toBe(true);
      expect(fake.tableNames()).not.toContain("HostSales");
    });

    it("ExcelApi 1.2 precheck false never calls convertToRange", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      const f = installTableUnlistExcel({ excelApi12: false });
      const result = await new OfficeJsAdapter().unlistTable("Sheet1", "Sales");
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(true);
        expect(result.evidence).toMatch(/ExcelApi 1\.2|convertToRange/);
      }
      expect(f.convertCalls()).toBe(0);
    });

    it("missing isSetSupported is unsupported with convert 0", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      const f = installTableUnlistExcel({ missingIsSetSupported: true });
      const result = await new OfficeJsAdapter().unlistTable("Sheet1", "Sales");
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.unsupported).toBe(true);
      expect(f.convertCalls()).toBe(0);
    });

    it("isSetSupported throw is unsupported with convert 0", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      const f = installTableUnlistExcel({ isSetSupportedThrows: true });
      const result = await new OfficeJsAdapter().unlistTable("Sheet1", "Sales");
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.unsupported).toBe(true);
      expect(f.convertCalls()).toBe(0);
    });

    it("missing convertToRange fails and is not a success", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      const f = installTableUnlistExcel({ hasConvertToRange: false });
      const result = await new OfficeJsAdapter().unlistTable("Sheet1", "Sales");
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.reason ?? "").toMatch(/convertToRange/);
      }
      expect(f.convertCalls()).toBe(0);
    });

    it("table.delete still hard-deletes (regression)", async () => {
      const adapter = new OfficeJsAdapter();
      const deleted = await adapter.deleteTable("Sheet1", "Sales");
      expect(deleted.ok).toBe(true);
      expect(fake.tableNames()).toHaveLength(0);
      expect(fake.convertCalls()).toBe(0);
    });
  });

  describe("executor + schema", () => {
    it("registers table.unlist as moderate with additionalProperties false", () => {
      const def = TOOL_DEFINITIONS.find((t) => t.name === "table.unlist");
      expect(def).toBeDefined();
      expect(def?.riskLevel).toBe("moderate");
      expect(def?.parameters).toMatchObject({
        additionalProperties: false,
        required: ["sheetName", "tableName"],
      });
    });

    it("MockHost parity: unlist keeps cells, removes table", async () => {
      const host = new MockHostAdapter();
      await host.writeRange("Sheet1", "A1:B2", [
        ["H1", "H2"],
        ["v1", "v2"],
      ]);
      await host.createTable({
        sheetName: "Sheet1",
        address: "A1:B2",
        name: "Sales",
      });
      const executor = new ToolExecutor(host);
      const result = await executor.execute({
        name: "table.unlist",
        arguments: { sheetName: "  Sheet1  ", tableName: "  Sales  " },
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data).toMatchObject({
          sheetName: "Sheet1",
          tableName: "Sales",
          unlisted: true,
        });
      }
      const listed = await host.listTables("Sheet1");
      expect(listed.ok && listed.data).toEqual([]);
      const cells = await host.readRange("Sheet1", "A1:B2");
      expect(cells.ok && cells.data.values).toEqual([
        ["H1", "H2"],
        ["v1", "v2"],
      ]);
    });

    it("rejects invalid args", async () => {
      const host = new MockHostAdapter();
      await host.createTable({ sheetName: "Sheet1", address: "A1:B2", name: "Sales" });
      const executor = new ToolExecutor(host);
      for (const args of [
        { sheetName: "Sheet1" },
        { tableName: "Sales" },
        { sheetName: "", tableName: "Sales" },
        { sheetName: "   ", tableName: "Sales" },
        { sheetName: "Sheet1", tableName: "" },
        { sheetName: "Sheet1", tableName: "   " },
        { sheetName: "Sheet1", tableName: "Sales", extra: 1 },
        { sheetName: 1, tableName: "Sales" },
        { sheetName: "Sheet1", tableName: false },
        { sheetName: null, tableName: "Sales" },
        { sheetName: "Sheet1", tableName: null },
        { sheetName: undefined, tableName: "Sales" },
        { sheetName: "Sheet1", tableName: undefined },
      ]) {
        const result = await executor.execute({
          name: "table.unlist",
          arguments: args as Record<string, unknown>,
        });
        expect(result.ok).toBe(false);
      }
    });

    it("table.delete remains hard-delete via executor", async () => {
      const host = new MockHostAdapter();
      await host.createTable({ sheetName: "Sheet1", address: "A1:B2", name: "Sales" });
      const executor = new ToolExecutor(host);
      const del = await executor.execute({
        name: "table.delete",
        arguments: { sheetName: "Sheet1", tableName: "Sales" },
      });
      expect(del.ok).toBe(true);
      const listed = await host.listTables();
      expect(listed.ok && listed.data).toEqual([]);
    });
  });

  describe("WPS", () => {
    beforeEach(() => {
      (globalThis as unknown as { window: unknown }).window = globalThis;
      (globalThis as unknown as { Application: unknown }).Application = {
        Name: "WPS",
        ActiveWorkbook: { Name: "Book1.xlsx" },
      };
    });
    afterEach(() => {
      delete (globalThis as { Application?: unknown }).Application;
    });

    it("returns typed unsupported for table.unlist", async () => {
      const result = await new WpsJsaAdapter().unlistTable("Sheet1", "Sales");
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(true);
        expect(result.host).toBe("wps-jsa");
        expect(result.capability).toBe("table.unlist");
        expect(result.evidence).toMatch(/ListObjects|Unlist|convertToRange/i);
      }
    });
  });
});
