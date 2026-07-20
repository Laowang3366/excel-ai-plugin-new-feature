import { afterEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { TOOL_DEFINITION_MAP, ToolExecutor } from "../shared/tools";
import { listChatReadOnlyTools } from "../shared/agentChat";
import {
  buildSheetOrder,
  sortNamedRanges,
  sortTables,
} from "../shared/host/workbookObjectsHelpers";
import { MockHostAdapter } from "./mockHost";
import { installWorkbookObjectsExcel } from "./fakes/officeJsWorkbookObjectsFake";

describe("workbook.objects.inspect", () => {
  describe("definition / validation / read-only", () => {
    it("is safe and registered with closed schema", () => {
      const def = TOOL_DEFINITION_MAP["workbook.objects.inspect"];
      expect(def).toBeDefined();
      expect(def.riskLevel).toBe("safe");
      expect(def.parameters).toMatchObject({
        type: "object",
        additionalProperties: false,
      });
      expect(listChatReadOnlyTools().some((t) => t.name === "workbook.objects.inspect")).toBe(
        true,
      );
    });

    it("rejects unknown args, non-integer max, and out-of-range max", async () => {
      const executor = new ToolExecutor(new MockHostAdapter());
      const unknown = await executor.execute({
        name: "workbook.objects.inspect",
        arguments: { extra: 1 },
      });
      expect(unknown.ok).toBe(false);
      if (!unknown.ok) expect(unknown.error).toMatch(/unknown field|extra/i);

      const nonInt = await executor.execute({
        name: "workbook.objects.inspect",
        arguments: { maxItemsPerCategory: 1.5 },
      });
      expect(nonInt.ok).toBe(false);
      if (!nonInt.ok) expect(nonInt.error).toMatch(/maxItemsPerCategory/);

      const low = await executor.execute({
        name: "workbook.objects.inspect",
        arguments: { maxItemsPerCategory: 0 },
      });
      expect(low.ok).toBe(false);

      const high = await executor.execute({
        name: "workbook.objects.inspect",
        arguments: { maxItemsPerCategory: 501 },
      });
      expect(high.ok).toBe(false);
    });
  });

  describe("position-first sorting helpers", () => {
    it("orders sheet objects by position before name, even when alphabet disagrees", () => {
      const order = buildSheetOrder([
        { name: "Sheet1", index: 0, isActive: true },
        { name: "Data", index: 1, isActive: false },
      ]);
      const tables = sortTables(
        [
          { name: "T_Data", sheetName: "Data", address: "A1", hasHeaders: true },
          { name: "T_B", sheetName: "Sheet1", address: "A1", hasHeaders: true },
          { name: "T_A", sheetName: "Sheet1", address: "B1", hasHeaders: true },
        ],
        order,
      );
      expect(tables.map((t) => `${t.sheetName}:${t.name}`)).toEqual([
        "Sheet1:T_A",
        "Sheet1:T_B",
        "Data:T_Data",
      ]);

      const names = sortNamedRanges(
        [
          { name: "Zed", refersTo: "=1", scope: "workbook" },
          { name: "DataLocal", refersTo: "=1", scope: "worksheet", sheetName: "Data" },
          { name: "Local", refersTo: "=1", scope: "worksheet", sheetName: "Sheet1" },
        ],
        order,
      );
      expect(names.map((n) => `${n.scope}:${n.sheetName ?? ""}:${n.name}`)).toEqual([
        "workbook::Zed",
        "worksheet:Sheet1:Local",
        "worksheet:Data:DataLocal",
      ]);
    });
  });

  describe("MockHost + executor", () => {
    it("aggregates sorted categories and truncates with real totalCount", async () => {
      const host = new MockHostAdapter();
      host.sheets = [
        { name: "Sheet1", index: 0, isActive: true },
        { name: "Data", index: 1, isActive: false },
      ];
      host.tables = [
        { name: "T_B", sheetName: "Sheet1", address: "A1:B2", hasHeaders: true },
        { name: "T_A", sheetName: "Sheet1", address: "C1:D2", hasHeaders: true },
        { name: "T_Data", sheetName: "Data", address: "A1:C3", hasHeaders: true },
      ];
      host.charts = [
        {
          name: "ChartB",
          sheetName: "Sheet1",
          chartType: "line",
          title: "B",
          left: 0,
          top: 0,
          width: 1,
          height: 1,
        },
        {
          name: "ChartA",
          sheetName: "Sheet1",
          chartType: "column",
          title: "A",
          left: 0,
          top: 0,
          width: 1,
          height: 1,
        },
      ];
      host.shapes = [
        {
          name: "ShapeB",
          sheetName: "Sheet1",
          type: "GeometricShape",
          geometricShapeType: "Ellipse",
          left: 0,
          top: 0,
          width: 1,
          height: 1,
        },
        {
          name: "ShapeA",
          sheetName: "Sheet1",
          type: "GeometricShape",
          geometricShapeType: "Rectangle",
          left: 0,
          top: 0,
          width: 1,
          height: 1,
        },
      ];
      await host.createNamedRange({
        name: "Zed",
        refersTo: "=Sheet1!$A$1",
        scope: "workbook",
      });
      await host.createNamedRange({
        name: "Local",
        refersTo: "=Data!$A$1",
        scope: "worksheet",
        sheetName: "Data",
      });

      const executor = new ToolExecutor(host);
      const full = await executor.execute({
        name: "workbook.objects.inspect",
        arguments: {},
      });
      expect(full.ok).toBe(true);
      if (!full.ok) return;
      const data = full.data as {
        tables: {
          items: { name: string; sheetName: string }[];
          totalCount: number;
          truncated: boolean;
          status: string;
        };
        charts: { items: { name: string }[]; totalCount: number };
        namedRanges: {
          items: { name: string; scope: string; sheetName?: string }[];
          status: string;
        };
        shapes: { items: { name: string }[] };
        sheets: { name: string }[];
        limitations: string[];
      };
      expect(data.tables.status).toBe("available");
      // Sheet1 (pos 0) before Data (pos 1) despite "Data" < "Sheet1" alphabetically.
      expect(data.tables.items.map((t) => `${t.sheetName}:${t.name}`)).toEqual([
        "Sheet1:T_A",
        "Sheet1:T_B",
        "Data:T_Data",
      ]);
      expect(data.charts.items.map((c) => c.name)).toEqual(["ChartA", "ChartB"]);
      expect(data.shapes.items.map((s) => s.name)).toEqual(["ShapeA", "ShapeB"]);
      expect(data.namedRanges.items.some((n) => n.scope === "worksheet" && n.sheetName === "Data")).toBe(
        true,
      );
      expect(data.limitations.length).toBeGreaterThan(0);

      const capped = await executor.execute({
        name: "workbook.objects.inspect",
        arguments: { maxItemsPerCategory: 1 },
      });
      expect(capped.ok).toBe(true);
      if (!capped.ok) return;
      const c = capped.data as {
        tables: { items: { name: string; sheetName: string }[]; totalCount: number; truncated: boolean };
      };
      expect(c.tables.items).toHaveLength(1);
      expect(c.tables.totalCount).toBe(3);
      expect(c.tables.truncated).toBe(true);
      // First after position sort is Sheet1:T_A
      expect(`${c.tables.items[0]!.sheetName}:${c.tables.items[0]!.name}`).toBe("Sheet1:T_A");
    });

    it("keeps other categories when one is unsupported or failed", async () => {
      const host = new MockHostAdapter();
      host.tables = [
        { name: "T1", sheetName: "Sheet1", address: "A1", hasHeaders: true },
      ];
      host.objectCategoryOverride = { charts: "unsupported", shapes: "failed" };
      const executor = new ToolExecutor(host);
      const result = await executor.execute({
        name: "workbook.objects.inspect",
        arguments: {},
      });
      expect(result.ok).toBe(true);
      if (!result.ok) return;
      const data = result.data as {
        tables: { status: string; items: unknown[] };
        charts: { status: string; items: unknown[]; reason?: string };
        shapes: { status: string; items: unknown[]; reason?: string };
        namedRanges: { status: string };
      };
      expect(data.tables.status).toBe("available");
      expect(data.tables.items).toHaveLength(1);
      expect(data.charts.status).toBe("unsupported");
      expect(data.charts.items).toEqual([]);
      expect(data.charts.reason).toMatch(/unsupported/i);
      expect(data.shapes.status).toBe("failed");
      expect(data.shapes.items).toEqual([]);
      expect(data.namedRanges.status).toBe("available");
    });

    it("fails entirely when base workbook inspect is forced to fail", async () => {
      const host = new MockHostAdapter();
      host.failCapability = "workbook.objects.inspect";
      const executor = new ToolExecutor(host);
      const result = await executor.execute({
        name: "workbook.objects.inspect",
        arguments: {},
      });
      expect(result.ok).toBe(false);
    });
  });

  describe("Office.js host", () => {
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
    });

    it("lists multi-sheet objects with position-first sort and bounded Excel.run count", async () => {
      const fake = installWorkbookObjectsExcel({ extraSheetCount: 5 });
      const adapter = new OfficeJsAdapter();
      const result = await adapter.inspectWorkbookObjects({ maxItemsPerCategory: 100 });
      expect(result.ok).toBe(true);
      if (!result.ok) return;
      expect(result.data.workbookName).toBe("Book1.xlsx");
      expect(result.data.sheets[0]?.name).toBe("Sheet1");
      expect(result.data.sheets[1]?.name).toBe("Data");
      expect(result.data.tables.status).toBe("available");
      // Position order: Sheet1 (0) before Data (1)
      expect(result.data.tables.items.map((t) => `${t.sheetName}:${t.name}`).slice(0, 3)).toEqual([
        "Sheet1:T_A",
        "Sheet1:T_B",
        "Data:T_Data",
      ]);
      expect(result.data.charts.items.map((c) => c.name)).toEqual(["ChartA", "ChartB"]);
      expect(result.data.shapes.items.map((s) => `${s.sheetName}:${s.name}`).slice(0, 3)).toEqual([
        "Sheet1:ShapeA",
        "Sheet1:ShapeB",
        "Data:DataShape",
      ]);
      const names = result.data.namedRanges.items;
      expect(names.some((n) => n.scope === "workbook" && n.name === "Alpha")).toBe(true);
      expect(names.some((n) => n.scope === "worksheet" && n.sheetName === "Data")).toBe(true);
      // Fixed bound: 1 baseline + 4 categories = 5, independent of sheet count (7 sheets here)
      expect(fake.sheetCount()).toBe(7);
      expect(fake.getRunCount()).toBe(5);
    });

    it("tables sync-time failure fails only tables; other categories stay available", async () => {
      const fake = installWorkbookObjectsExcel({ failTablesOnSync: true });
      const adapter = new OfficeJsAdapter();
      const result = await adapter.inspectWorkbookObjects({});
      expect(result.ok).toBe(true);
      if (!result.ok) return;
      expect(result.data.tables.status).toBe("failed");
      expect(result.data.tables.items).toEqual([]);
      expect(result.data.tables.reason).toMatch(/tables sync failed/i);
      expect(result.data.charts.status).toBe("available");
      expect(result.data.charts.items.length).toBeGreaterThan(0);
      expect(result.data.namedRanges.status).toBe("available");
      expect(result.data.shapes.status).toBe("available");
      // Still one run per phase even when tables fail
      expect(fake.getRunCount()).toBe(5);
    });

    it("fails whole tool when workbook baseline sync fails", async () => {
      const fake = installWorkbookObjectsExcel({ failWorkbookOnSync: true });
      const adapter = new OfficeJsAdapter();
      const result = await adapter.inspectWorkbookObjects({});
      expect(result.ok).toBe(false);
      // Only baseline run attempted
      expect(fake.getRunCount()).toBe(1);
    });
  });

  describe("WPS JSA host", () => {
    afterEach(() => {
      delete (globalThis as { Application?: unknown }).Application;
      delete (globalThis as { window?: { Application?: unknown } }).window?.Application;
    });

    function installWps(options?: {
      withWorkbookNames?: boolean;
      withWorksheetNames?: boolean;
      workbookNamesEmpty?: boolean;
    }) {
      const withWorkbookNames = options?.withWorkbookNames ?? true;
      const withWorksheetNames = options?.withWorksheetNames ?? true;
      const sheets = [
        { Name: "Sheet1", Index: 1 },
        { Name: "Data", Index: 2 },
      ];
      const wbNames: Array<{ Name: string; RefersTo: string; Visible: boolean }> =
        options?.workbookNamesEmpty
          ? []
          : [{ Name: "Wb1", RefersTo: "=Sheet1!$A$1", Visible: true }];
      const sheetNames: Record<string, Array<{ Name: string; RefersTo: string; Visible: boolean }>> =
        {
          Sheet1: [{ Name: "Local1", RefersTo: "=Sheet1!$B$1", Visible: true }],
          Data: [],
        };

      function namesApi(list: Array<{ Name: string; RefersTo: string; Visible: boolean }>) {
        return {
          get Count() {
            return list.length;
          },
          Item(i: number) {
            const rec = list[i - 1];
            if (!rec) throw new Error("missing");
            return {
              get Name() {
                return rec.Name;
              },
              get RefersTo() {
                return rec.RefersTo;
              },
              get Visible() {
                return rec.Visible;
              },
              Delete() {},
            };
          },
          Add() {
            throw new Error("not used");
          },
        };
      }

      function sheetObj(s: { Name: string; Index: number }) {
        return {
          Name: s.Name,
          Index: s.Index,
          Names: withWorksheetNames ? namesApi(sheetNames[s.Name] ?? []) : undefined,
        };
      }

      const workbook = {
        Name: "BookWps.xlsx",
        get ActiveSheet() {
          return {
            Name: "Sheet1",
            UsedRange: { Address: "A1:B2" },
            Names: withWorksheetNames ? namesApi(sheetNames.Sheet1 ?? []) : undefined,
          };
        },
        Worksheets: {
          get Count() {
            return sheets.length;
          },
          Item(indexOrName: number | string) {
            if (typeof indexOrName === "number") {
              const s = sheets[indexOrName - 1];
              if (!s) throw new Error("sheet");
              return sheetObj(s);
            }
            const s = sheets.find((x) => x.Name === indexOrName);
            if (!s) throw new Error("sheet");
            return sheetObj(s);
          },
        },
        Names: withWorkbookNames ? namesApi(wbNames) : undefined,
      };

      const Application = {
        ActiveWorkbook: workbook,
        Workbooks: { Count: 1, Item: () => workbook },
      };
      (globalThis as unknown as { Application: unknown }).Application = Application;
      (globalThis as unknown as { window: { Application: unknown } }).window = {
        Application,
      };
    }

    it("returns sheets+names available and table/chart/shape unsupported", async () => {
      installWps();
      const adapter = new WpsJsaAdapter();
      const result = await adapter.inspectWorkbookObjects({});
      expect(result.ok).toBe(true);
      if (!result.ok) return;
      expect(result.data.workbookName).toBe("BookWps.xlsx");
      expect(result.data.sheets.length).toBeGreaterThan(0);
      expect(result.data.namedRanges.status).toBe("available");
      expect(result.data.namedRanges.items.some((n) => n.scope === "workbook")).toBe(true);
      expect(result.data.tables.status).toBe("unsupported");
      expect(result.data.tables.items).toEqual([]);
      expect(result.data.tables.reason).toBeTruthy();
      expect(result.data.charts.status).toBe("unsupported");
      expect(result.data.shapes.status).toBe("unsupported");
    });

    it("namedRanges is unsupported when all Names members are missing (not available+[])", async () => {
      installWps({ withWorkbookNames: false, withWorksheetNames: false });
      const adapter = new WpsJsaAdapter();
      const result = await adapter.inspectWorkbookObjects({});
      expect(result.ok).toBe(true);
      if (!result.ok) return;
      expect(result.data.namedRanges.status).toBe("unsupported");
      expect(result.data.namedRanges.items).toEqual([]);
      expect(result.data.namedRanges.totalCount).toBeNull();
      expect(result.data.namedRanges.reason).toMatch(/Names/i);
      expect(result.data.namedRanges.evidence).toBeTruthy();
    });

    it("namedRanges stays available when workbook Names succeeds empty and worksheet Names unsupported", async () => {
      installWps({
        withWorkbookNames: true,
        workbookNamesEmpty: true,
        withWorksheetNames: false,
      });
      const adapter = new WpsJsaAdapter();
      const result = await adapter.inspectWorkbookObjects({});
      expect(result.ok).toBe(true);
      if (!result.ok) return;
      expect(result.data.namedRanges.status).toBe("available");
      expect(result.data.namedRanges.totalCount).toBe(0);
      expect(result.data.namedRanges.items).toEqual([]);
      const lim = [
        ...(result.data.limitations ?? []),
        ...(result.data.namedRanges.limitations ?? []),
      ].join(" ");
      expect(lim).toMatch(/worksheet\.Names|unavailable/i);
    });
  });
});
