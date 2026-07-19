import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import type { CellValue } from "../shared/host/types";

type CellState = { values: CellValue[][]; formulas: string[][] };

function cellKey(address: string): string {
  return address.toUpperCase();
}

function installFakeExcel() {
  const sheets = new Map<
    string,
    { name: string; position: number; cells: Map<string, CellState> }
  >();
  sheets.set("Sheet1", { name: "Sheet1", position: 0, cells: new Map() });

  function getSheet(name: string) {
    const sheet = sheets.get(name);
    if (!sheet) throw new Error(`Sheet not found: ${name}`);
    return sheet;
  }

  function makeRange(sheetName: string, address: string) {
    const sheet = getSheet(sheetName);
    const key = cellKey(address);
    const range: {
      address: string;
      values: CellValue[][];
      formulas: string[][];
      load: (props: string) => void;
      clear: () => void;
    } = {
      address: `${sheetName}!${address}`,
      get values() {
        return sheet.cells.get(key)?.values ?? [[null]];
      },
      set values(next: CellValue[][]) {
        const prev = sheet.cells.get(key);
        sheet.cells.set(key, {
          values: next,
          formulas: prev?.formulas ?? next.map((row) => row.map(() => "")),
        });
      },
      get formulas() {
        return sheet.cells.get(key)?.formulas ?? [[""]];
      },
      set formulas(next: string[][]) {
        const prev = sheet.cells.get(key);
        sheet.cells.set(key, {
          values: prev?.values ?? next.map((row) => row.map(() => null)),
          formulas: next,
        });
      },
      load() {},
      clear() {
        sheet.cells.delete(key);
      },
    };
    return range;
  }

  const context = {
    workbook: {
      name: "Book1.xlsx",
      load() {},
      worksheets: {
        items: [] as { name: string; position: number }[],
        load() {
          context.workbook.worksheets.items = [...sheets.values()].map((sheet) => ({
            name: sheet.name,
            position: sheet.position,
          }));
        },
        getActiveWorksheet() {
          const sheet = [...sheets.values()][0]!;
          return {
            name: sheet.name,
            position: sheet.position,
            load() {},
            getRange(address: string) {
              return makeRange(sheet.name, address);
            },
            delete() {
              sheets.delete(sheet.name);
            },
          };
        },
        getItem(name: string) {
          const sheet = getSheet(name);
          return {
            get name() {
              return sheet.name;
            },
            set name(next: string) {
              sheets.delete(sheet.name);
              sheet.name = next;
              sheets.set(next, sheet);
            },
            position: sheet.position,
            load() {},
            getRange(address: string) {
              return makeRange(sheet.name, address);
            },
            delete() {
              sheets.delete(sheet.name);
            },
          };
        },
        add(name?: string) {
          const sheetName = name ?? `Sheet${sheets.size + 1}`;
          const sheet = { name: sheetName, position: sheets.size, cells: new Map<string, CellState>() };
          sheets.set(sheetName, sheet);
          return {
            name: sheet.name,
            position: sheet.position,
            load() {},
            getRange(address: string) {
              return makeRange(sheet.name, address);
            },
            delete() {
              sheets.delete(sheet.name);
            },
          };
        },
      },
      getSelectedRange() {
        return makeRange("Sheet1", "A1");
      },
    },
    async sync() {},
  };

  (globalThis as unknown as { window: unknown }).window = globalThis;
  (globalThis as unknown as { Excel: { run: typeof context extends never ? never : Function } }).Excel =
    {
      run: async <T>(fn: (ctx: typeof context) => Promise<T>) => fn(context),
    };

  return { sheets };
}

function installFakeWps() {
  const sheets = new Map<
    string,
    { Name: string; Index: number; cells: Map<string, CellState> }
  >();
  sheets.set("Sheet1", { Name: "Sheet1", Index: 1, cells: new Map() });

  function sheetApi(name: string) {
    const sheet = sheets.get(name);
    if (!sheet) throw new Error(`missing ${name}`);
    return {
      get Name() {
        return sheet.Name;
      },
      set Name(next: string) {
        sheets.delete(sheet.Name);
        sheet.Name = next;
        sheets.set(next, sheet);
      },
      Index: sheet.Index,
      Range(address: string) {
        const key = cellKey(address);
        const range = {
          Address: `${sheet.Name}!${address}`,
          get Value2() {
            return sheet.cells.get(key)?.values ?? [[null]];
          },
          set Value2(next: CellValue[][]) {
            const prev = sheet.cells.get(key);
            sheet.cells.set(key, {
              values: next,
              formulas: prev?.formulas ?? next.map((row) => row.map(() => "")),
            });
          },
          get Formula() {
            const formulas = sheet.cells.get(key)?.formulas ?? [[""]];
            return formulas.length === 1 && formulas[0]?.length === 1
              ? formulas[0][0]
              : formulas;
          },
          set Formula(next: string | string[][]) {
            const matrix = typeof next === "string" ? [[next]] : next;
            const prev = sheet.cells.get(key);
            sheet.cells.set(key, {
              values: prev?.values ?? matrix.map((row) => row.map(() => null)),
              formulas: matrix,
            });
          },
          Clear() {
            sheet.cells.delete(key);
          },
        };
        return range;
      },
      Delete() {
        sheets.delete(sheet.Name);
      },
    };
  }

  const workbook = {
    Name: "Book1.xlsx",
    get ActiveSheet() {
      return sheetApi([...sheets.keys()][0]!);
    },
    Worksheets: {
      get Count() {
        return sheets.size;
      },
      Item(indexOrName: number | string) {
        if (typeof indexOrName === "number") {
          const name = [...sheets.keys()][indexOrName - 1];
          if (!name) throw new Error("index");
          return sheetApi(name);
        }
        return sheetApi(indexOrName);
      },
      Add() {
        const name = `Sheet${sheets.size + 1}`;
        sheets.set(name, { Name: name, Index: sheets.size + 1, cells: new Map() });
        return sheetApi(name);
      },
    },
  };

  (globalThis as unknown as { window: unknown }).window = globalThis;
  (globalThis as unknown as { Application: unknown }).Application = {
    Name: "WPS 表格",
    ActiveWorkbook: workbook,
    Selection: workbook.ActiveSheet.Range("A1"),
  };

  return { sheets };
}

describe("OfficeJsAdapter with fake Excel.run", () => {
  beforeEach(() => {
    installFakeExcel();
  });

  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  it("reads/writes range and formulas via formulas property", async () => {
    const adapter = new OfficeJsAdapter();
    const written = await adapter.writeFormulas("Sheet1", "B2", [["=1+2"]]);
    expect(written.ok).toBe(true);
    if (!written.ok) return;
    expect(written.data.formulas[0][0]).toBe("=1+2");

    const valuesWrite = await adapter.writeRange("Sheet1", "A1", [["x"]]);
    expect(valuesWrite.ok).toBe(true);

    const read = await adapter.readRange("Sheet1", "B2");
    expect(read.ok).toBe(true);
    if (read.ok) expect(read.data.formulas[0][0]).toBe("=1+2");
  });

  it("lists/adds/renames/deletes sheets", async () => {
    const adapter = new OfficeJsAdapter();
    const added = await adapter.addSheet("Data");
    expect(added.ok).toBe(true);
    const renamed = await adapter.renameSheet("Data", "Metrics");
    expect(renamed.ok).toBe(true);
    const listed = await adapter.listSheets();
    expect(listed.ok).toBe(true);
    if (listed.ok) {
      expect(listed.data.map((s) => s.name)).toContain("Metrics");
    }
    const deleted = await adapter.deleteSheet("Metrics");
    expect(deleted.ok).toBe(true);
  });
});

describe("WpsJsaAdapter with fake Application", () => {
  beforeEach(() => {
    installFakeWps();
  });

  afterEach(() => {
    delete (globalThis as { Application?: unknown }).Application;
  });

  it("writes Formula property not Value2 for formulas", async () => {
    const adapter = new WpsJsaAdapter();
    const written = await adapter.writeFormulas("Sheet1", "C1", [["=SUM(1,2)"]]);
    expect(written.ok).toBe(true);
    if (!written.ok) return;
    expect(written.data.formulas[0][0]).toBe("=SUM(1,2)");

    const read = await adapter.readRange("Sheet1", "C1", "none");
    expect(read.ok).toBe(true);
    if (read.ok) expect(read.data.formulas[0][0]).toBe("=SUM(1,2)");
  });

  it("supports sheet list/add/rename/delete", async () => {
    const adapter = new WpsJsaAdapter();
    expect((await adapter.addSheet("Tmp")).ok).toBe(true);
    expect((await adapter.renameSheet("Tmp", "Final")).ok).toBe(true);
    const listed = await adapter.listSheets();
    expect(listed.ok).toBe(true);
    if (listed.ok) expect(listed.data.map((s) => s.name)).toContain("Final");
    expect((await adapter.deleteSheet("Final")).ok).toBe(true);
  });
});
