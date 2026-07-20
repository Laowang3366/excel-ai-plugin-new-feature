import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { installTableFilterExcel } from "./fakes/officeJsTableFilterFake";
import { installTableSortExcel } from "./fakes/officeJsTableSortFake";
import { MockHostAdapter } from "./mockHost";

describe("phase39 table filter/sort + first/last column", () => {
  describe("schema", () => {
    it("registers filter/sort tools with closed schemas", () => {
      const names = TOOL_DEFINITIONS.map((t) => t.name);
      for (const name of [
        "table.filter.get",
        "table.filter.apply",
        "table.filter.clear",
        "table.sort.get",
        "table.sort.apply",
        "table.sort.clear",
      ]) {
        expect(names).toContain(name);
        const def = TOOL_DEFINITIONS.find((t) => t.name === name);
        expect(def?.parameters.additionalProperties).toBe(false);
      }
      const update = TOOL_DEFINITIONS.find((t) => t.name === "table.update");
      expect(update?.parameters.properties).toMatchObject({
        showFirstColumn: { type: "boolean" },
        showLastColumn: { type: "boolean" },
      });
      const apply = TOOL_DEFINITIONS.find((t) => t.name === "table.filter.apply");
      expect(apply?.riskLevel).toBe("moderate");
      const applyProps = apply?.parameters.properties as
        | Record<string, { enum?: string[] }>
        | undefined;
      expect(applyProps?.filterOn).toMatchObject({
        enum: expect.arrayContaining([
          "values",
          "custom",
          "topItems",
          "bottomItems",
          "topPercent",
          "bottomPercent",
        ]),
      });
    });
  });

  describe("Office.js table.filter", () => {
    let fake: ReturnType<typeof installTableFilterExcel>;
    beforeEach(() => {
      fake = installTableFilterExcel({
        hostSheetName: "HostSheet",
        hostTableName: "HostSales",
      });
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
    });

    it("apply values converts 1-based columnIndex and returns host names", async () => {
      const result = await new OfficeJsAdapter().applyTableFilter({
        sheetName: "Sheet1",
        tableName: "Sales",
        columnIndex: 2,
        filterOn: "values",
        values: ["East", "West"],
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.sheetName).toBe("HostSheet");
        expect(result.data.tableName).toBe("HostSales");
        expect(result.data.enabled).toBe(true);
        expect(result.data.columnIndex).toBe(2);
        expect(result.data.filterOn).toBe("values");
      }
      expect(fake.applyCalls()).toBe(1);
      expect(fake.lastApply()?.columnIndex).toBe(1); // 0-based
      expect(fake.lastApply()?.criteria).toMatchObject({
        filterOn: "Values",
        values: ["East", "West"],
      });
    });

    it("apply custom builds criterion + operator", async () => {
      const result = await new OfficeJsAdapter().applyTableFilter({
        sheetName: "Sheet1",
        tableName: "Sales",
        columnIndex: 1,
        filterOn: "custom",
        criterion1: ">10",
        criterion2: "<100",
        operator: "and",
      });
      expect(result.ok).toBe(true);
      expect(fake.lastApply()?.criteria).toMatchObject({
        filterOn: "Custom",
        criterion1: ">10",
        criterion2: "<100",
        operator: "And",
      });
    });

    it("apply topItems uses threshold as criterion1", async () => {
      const result = await new OfficeJsAdapter().applyTableFilter({
        sheetName: "Sheet1",
        tableName: "Sales",
        columnIndex: 3,
        filterOn: "topItems",
        threshold: 5,
      });
      expect(result.ok).toBe(true);
      expect(fake.lastApply()?.criteria).toMatchObject({
        filterOn: "TopItems",
        criterion1: "5",
      });
    });

    it("clear clears criteria", async () => {
      await new OfficeJsAdapter().applyTableFilter({
        sheetName: "Sheet1",
        tableName: "Sales",
        columnIndex: 1,
        filterOn: "values",
        values: ["A"],
      });
      const cleared = await new OfficeJsAdapter().clearTableFilter({
        sheetName: "Sheet1",
        tableName: "Sales",
      });
      expect(cleared.ok).toBe(true);
      if (cleared.ok) expect(cleared.data.enabled).toBe(false);
      expect(fake.clearCalls()).toBe(1);
    });

    it("get reads enabled via ExcelApi 1.9", async () => {
      await new OfficeJsAdapter().applyTableFilter({
        sheetName: "Sheet1",
        tableName: "Sales",
        columnIndex: 1,
        filterOn: "values",
        values: ["A"],
      });
      const got = await new OfficeJsAdapter().getTableFilter({
        sheetName: "Sheet1",
        tableName: "Sales",
      });
      expect(got.ok).toBe(true);
      if (got.ok) {
        expect(got.data.enabled).toBe(true);
        expect(got.data.sheetName).toBe("HostSheet");
      }
    });

    it("ExcelApi 1.2 false never calls apply", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      const f = installTableFilterExcel({ excelApi12: false });
      const result = await new OfficeJsAdapter().applyTableFilter({
        sheetName: "Sheet1",
        tableName: "Sales",
        columnIndex: 1,
        filterOn: "values",
        values: ["A"],
      });
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.unsupported).toBe(true);
      expect(f.applyCalls()).toBe(0);
    });

    it("ExcelApi 1.9 false makes get unsupported", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installTableFilterExcel({ excelApi19: false });
      const result = await new OfficeJsAdapter().getTableFilter({
        sheetName: "Sheet1",
        tableName: "Sales",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(true);
        expect(result.evidence ?? result.reason ?? "").toMatch(/1\.9|enabled/);
      }
    });

    it("missing apply is ordinary fail", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installTableFilterExcel({ hasApply: false });
      const result = await new OfficeJsAdapter().applyTableFilter({
        sheetName: "Sheet1",
        tableName: "Sales",
        columnIndex: 1,
        filterOn: "values",
        values: ["A"],
      });
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.unsupported).not.toBe(true);
    });
  });

  describe("Office.js table.sort", () => {
    let fake: ReturnType<typeof installTableSortExcel>;
    beforeEach(() => {
      fake = installTableSortExcel({
        hostSheetName: "HostSheet",
        hostTableName: "HostSales",
      });
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
    });

    it("apply converts 1-based fields and readbacks host fields", async () => {
      const result = await new OfficeJsAdapter().applyTableSort({
        sheetName: "Sheet1",
        tableName: "Sales",
        fields: [
          { columnIndex: 2, ascending: false },
          { columnIndex: 1, ascending: true },
        ],
        matchCase: true,
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.sheetName).toBe("HostSheet");
        expect(result.data.tableName).toBe("HostSales");
        expect(result.data.fields).toEqual([
          { columnIndex: 2, ascending: false },
          { columnIndex: 1, ascending: true },
        ]);
      }
      expect(fake.applyCalls()).toBe(1);
      expect(fake.fields()).toEqual([
        { key: 1, ascending: false },
        { key: 0, ascending: true },
      ]);
    });

    it("clear empties fields", async () => {
      await new OfficeJsAdapter().applyTableSort({
        sheetName: "Sheet1",
        tableName: "Sales",
        fields: [{ columnIndex: 1 }],
      });
      const cleared = await new OfficeJsAdapter().clearTableSort({
        sheetName: "Sheet1",
        tableName: "Sales",
      });
      expect(cleared.ok).toBe(true);
      if (cleared.ok) expect(cleared.data.fields).toEqual([]);
      expect(fake.clearCalls()).toBe(1);
    });

    it("get returns applied fields", async () => {
      await new OfficeJsAdapter().applyTableSort({
        sheetName: "Sheet1",
        tableName: "Sales",
        fields: [{ columnIndex: 3, ascending: false }],
      });
      const got = await new OfficeJsAdapter().getTableSort({
        sheetName: "Sheet1",
        tableName: "Sales",
      });
      expect(got.ok).toBe(true);
      if (got.ok) {
        expect(got.data.fields).toEqual([{ columnIndex: 3, ascending: false }]);
      }
    });

    it("ExcelApi 1.2 false never calls apply", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      const f = installTableSortExcel({ excelApi12: false });
      const result = await new OfficeJsAdapter().applyTableSort({
        sheetName: "Sheet1",
        tableName: "Sales",
        fields: [{ columnIndex: 1 }],
      });
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.unsupported).toBe(true);
      expect(f.applyCalls()).toBe(0);
    });
  });

  describe("executor + mock host", () => {
    it("routes filter/sort tools and rejects unknown fields", async () => {
      const host = new MockHostAdapter();
      host.tables.push({
        name: "T1",
        sheetName: "Sheet1",
        address: "Sheet1!A1:C3",
        hasHeaders: true,
      });
      const executor = new ToolExecutor(host);

      const applied = await executor.execute({
        name: "table.filter.apply",
        arguments: {
          sheetName: "Sheet1",
          tableName: "T1",
          columnIndex: 1,
          filterOn: "values",
          values: ["x"],
        },
      });
      expect(applied.ok).toBe(true);

      const sorted = await executor.execute({
        name: "table.sort.apply",
        arguments: {
          sheetName: "Sheet1",
          tableName: "T1",
          fields: [{ columnIndex: 2, ascending: false }],
        },
      });
      expect(sorted.ok).toBe(true);
      if (sorted.ok) {
        expect(sorted.data).toMatchObject({
          fields: [{ columnIndex: 2, ascending: false }],
        });
      }

      const bad = await executor.execute({
        name: "table.filter.apply",
        arguments: {
          sheetName: "Sheet1",
          tableName: "T1",
          columnIndex: 1,
          filterOn: "values",
          values: ["x"],
          unknown: true,
        },
      });
      expect(bad.ok).toBe(false);
      if (!bad.ok) expect(bad.error).toMatch(/unknown field/);
    });

    it("table.update accepts showFirstColumn/showLastColumn", async () => {
      const host = new MockHostAdapter();
      host.tables.push({
        name: "T1",
        sheetName: "Sheet1",
        address: "Sheet1!A1:C3",
        hasHeaders: true,
      });
      const executor = new ToolExecutor(host);
      const result = await executor.execute({
        name: "table.update",
        arguments: {
          sheetName: "Sheet1",
          tableName: "T1",
          showFirstColumn: true,
          showLastColumn: false,
        },
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data).toMatchObject({
          showFirstColumn: true,
          showLastColumn: false,
        });
      }
    });

    it("rejects values mode without values", async () => {
      const host = new MockHostAdapter();
      host.tables.push({
        name: "T1",
        sheetName: "Sheet1",
        address: "Sheet1!A1:C3",
        hasHeaders: true,
      });
      const executor = new ToolExecutor(host);
      const result = await executor.execute({
        name: "table.filter.apply",
        arguments: {
          sheetName: "Sheet1",
          tableName: "T1",
          columnIndex: 1,
          filterOn: "values",
        },
      });
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.error).toMatch(/values/);
    });
  });

  describe("WPS", () => {
    it("returns typed unsupported for filter/sort tools", async () => {
      const executor = new ToolExecutor(new WpsJsaAdapter());
      for (const call of [
        { name: "table.filter.get" as const, arguments: { sheetName: "S", tableName: "T" } },
        {
          name: "table.filter.apply" as const,
          arguments: {
            sheetName: "S",
            tableName: "T",
            columnIndex: 1,
            filterOn: "values",
            values: ["a"],
          },
        },
        { name: "table.filter.clear" as const, arguments: { sheetName: "S", tableName: "T" } },
        { name: "table.sort.get" as const, arguments: { sheetName: "S", tableName: "T" } },
        {
          name: "table.sort.apply" as const,
          arguments: {
            sheetName: "S",
            tableName: "T",
            fields: [{ columnIndex: 1 }],
          },
        },
        { name: "table.sort.clear" as const, arguments: { sheetName: "S", tableName: "T" } },
      ]) {
        const result = await executor.execute(call);
        expect(result.ok).toBe(false);
        if (!result.ok) expect(result.unsupported).toBe(true);
      }
    });
  });
});
