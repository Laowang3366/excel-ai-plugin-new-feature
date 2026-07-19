import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { absoluteA1FromOrigin } from "../shared/host/a1Address";
import type { CellValue } from "../shared/host/types";

function installOfficeExpandFake(options?: { emptyUsedRange?: boolean }) {
  const cells = new Map<string, { values: CellValue[][]; formulas: string[][] }>();
  cells.set("A1", { values: [[1]], formulas: [["=1"]] });
  cells.set("B2", {
    values: [
      [10, 11],
      [12, 13],
    ],
    formulas: [
      ["=10", ""],
      ["", "=13"],
    ],
  });
  cells.set("D5", { values: [[5]], formulas: [["=5"]] });

  // Pending cell address loads — applied only on context.sync() (Office.js contract).
  const pendingCellLoads: Array<{ apply: () => void }> = [];
  let syncCount = 0;

  function makeCellProxy(a1: string) {
    let available = false;
    let wantsAddress = false;
    const proxy = {
      get address() {
        return available ? `Sheet1!${a1}` : "";
      },
      load(props?: string) {
        if (!props || props.includes("address")) {
          wantsAddress = true;
          pendingCellLoads.push({
            apply: () => {
              if (wantsAddress) available = true;
            },
          });
        }
      },
    };
    return proxy;
  }

  function makeRange(address: string) {
    const key = (address.includes("!") ? address.split("!")[1]! : address).toUpperCase();
    const origin = key.split(":")[0]!;
    let state = cells.get(origin) ?? cells.get(key) ?? {
      values: [[null as CellValue]],
      formulas: [[""]],
    };
    if (key.includes(":") && !cells.has(key)) {
      state = {
        values: [
          [1, 2],
          [3, 4],
        ],
        formulas: [
          ["=1", ""],
          ["", ""],
        ],
      };
      cells.set(key, state);
    }
    const range: Record<string, unknown> = {
      address: `Sheet1!${key}`,
      rowCount: state.values.length,
      columnCount: state.values[0]?.length ?? 1,
      get values() {
        return state.values;
      },
      get formulas() {
        return state.formulas;
      },
      load() {},
      clear() {},
      getSpillingToRange() {
        return makeRange("A1:A3");
      },
      getSurroundingRegion() {
        return makeRange("A1:C3");
      },
      getCurrentArray() {
        return makeRange("A1:B2");
      },
      getCell(row: number, col: number) {
        return makeCellProxy(absoluteA1FromOrigin(origin, row, col));
      },
    };
    return range;
  }

  let sheetPosition = 0;
  const sheetApi = {
    name: "Sheet1",
    load() {},
    getRange(address: string) {
      return makeRange(address);
    },
    getUsedRangeOrNullObject() {
      if (options?.emptyUsedRange) {
        return {
          isNullObject: true,
          address: "",
          values: [],
          formulas: [],
          load() {},
          getCell() {
            return { address: "", load() {} };
          },
        };
      }
      const r = makeRange("A1:B2") as {
        isNullObject?: boolean;
        load: (...args: string[]) => void;
      };
      r.isNullObject = false;
      return r;
    },
    delete() {},
    copy() {
      let copiedName = "Sheet1_Copy";
      return {
        get name() {
          return copiedName;
        },
        set name(v: string) {
          copiedName = v;
        },
        position: 1,
        load() {},
      };
    },
    get position() {
      return sheetPosition;
    },
    set position(v: number) {
      sheetPosition = v;
    },
    tables: { items: [], load() {}, add() {}, getItem() {} },
    charts: { items: [], load() {}, add() {}, getItem() {} },
  };

  const context = {
    workbook: {
      name: "Book1.xlsx",
      load() {},
      worksheets: {
        items: [sheetApi],
        load() {},
        getActiveWorksheet() {
          return sheetApi;
        },
        getItem() {
          return sheetApi;
        },
        add(name?: string) {
          return { name: name ?? "Sheet2", position: 1, load() {} };
        },
      },
      getSelectedRange() {
        return makeRange("A1");
      },
    },
    async sync() {
      syncCount += 1;
      // Apply all pending cell address loads (Office.js semantics).
      const batch = pendingCellLoads.splice(0, pendingCellLoads.length);
      for (const item of batch) item.apply();
    },
  };

  (globalThis as unknown as { __phase4SyncCount?: () => number }).__phase4SyncCount = () =>
    syncCount;

  (globalThis as unknown as { window: unknown }).window = globalThis;
  (globalThis as unknown as { Excel: { run: Function } }).Excel = {
    run: async <T>(fn: (ctx: typeof context) => Promise<T>) => fn(context),
  };
}

describe("phase4 Office.js expand/formula.context/sheet ops", () => {
  beforeEach(() => {
    installOfficeExpandFake();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  it("expands spill/currentArray/currentRegion", async () => {
    const adapter = new OfficeJsAdapter();
    for (const expand of ["spill", "currentArray", "currentRegion"] as const) {
      const result = await adapter.readRange("Sheet1", "A1", expand);
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.expandMode).toBe(expand);
        expect(result.data.expanded).toBe(true);
      }
    }
  });

  it("returns strict A1 addresses for formula.context from B2/D5 origins", async () => {
    const adapter = new OfficeJsAdapter();
    const syncBefore = (globalThis as { __phase4SyncCount?: () => number }).__phase4SyncCount?.() ?? 0;
    const fromB2 = await adapter.getFormulaContext("Sheet1", "B2");
    const syncAfter = (globalThis as { __phase4SyncCount?: () => number }).__phase4SyncCount?.() ?? 0;
    // At least: load range + batch cell addresses sync.
    expect(syncAfter).toBeGreaterThan(syncBefore);

    expect(fromB2.ok).toBe(true);
    if (fromB2.ok) {
      expect(fromB2.data.address).toBe("B2");
      const addrs = fromB2.data.formulas.map((f) => f.address);
      expect(addrs).toEqual(expect.arrayContaining(["B2", "C3"]));
      // Strict A1 only — no R1C1, no sheet prefix.
      for (const a of addrs) {
        expect(a).toMatch(/^[A-Z]+\d+$/);
        expect(a).not.toMatch(/^R\d+C\d+$/i);
        expect(a).not.toContain("!");
      }
    }

    const fromD5 = await adapter.getFormulaContext("Sheet1", "D5");
    expect(fromD5.ok).toBe(true);
    if (fromD5.ok) {
      expect(fromD5.data.formulas).toEqual([
        expect.objectContaining({ address: "D5", formula: "=5" }),
      ]);
    }

    const copied = await adapter.copySheet("Sheet1", "Copy1");
    expect(copied.ok).toBe(true);
    const moved = await adapter.moveSheet("Sheet1", 2);
    expect(moved.ok).toBe(true);
    if (moved.ok) expect(moved.data.index).toBe(2);
  });

  it("auto-spills single-cell read when expand is omitted", async () => {
    const adapter = new OfficeJsAdapter();
    const auto = await adapter.readRange("Sheet1", "A1");
    expect(auto.ok).toBe(true);
    if (auto.ok) {
      expect(auto.data.expandMode).toBe("spill");
      expect(auto.data.expanded).toBe(true);
    }
    const forced = await adapter.readRange("Sheet1", "A1", "none");
    expect(forced.ok).toBe(true);
    if (forced.ok) expect(forced.data.expandMode).toBe("none");
  });

  it("handles empty UsedRange for formula.context without loading isNullObject on normal range", async () => {
    delete (globalThis as { Excel?: unknown }).Excel;
    installOfficeExpandFake({ emptyUsedRange: true });
    const adapter = new OfficeJsAdapter();
    const empty = await adapter.getFormulaContext("Sheet1");
    expect(empty.ok).toBe(true);
    if (empty.ok) {
      expect(empty.data.formulas).toEqual([]);
      expect(empty.data.address).toBe("");
    }
    const specified = await adapter.getFormulaContext("Sheet1", "A1");
    expect(specified.ok).toBe(true);
    if (specified.ok) {
      expect(specified.data.formulas[0]?.address).toBe("A1");
    }
  });
});

describe("phase4 WPS expand/copy unsupported", () => {
  beforeEach(() => {
    (globalThis as unknown as { window: unknown }).window = globalThis;
    (globalThis as unknown as { Application: unknown }).Application = {
      Name: "WPS",
      ActiveWorkbook: {
        Name: "Book1.xlsx",
        ActiveSheet: {
          Name: "Sheet1",
          UsedRange: { Address: "B2:C3" },
          Range: (address: string) => ({
            Address: address,
            Value2: [
              [3, 4],
              [5, 6],
            ],
            Formula: [
              ["=1+2", ""],
              ["", "=6"],
            ],
          }),
        },
        Worksheets: {
          Count: 1,
          Item: () => ({
            Name: "Sheet1",
            Index: 1,
            UsedRange: { Address: "B2:C3" },
            Range: (address: string) => ({
              Address: address,
              Value2: [
                [3, 4],
                [5, 6],
              ],
              Formula: [
                ["=1+2", ""],
                ["", "=6"],
              ],
            }),
          }),
        },
      },
    };
  });
  afterEach(() => {
    delete (globalThis as { Application?: unknown }).Application;
  });

  it("rejects default single-cell spill and expand; formula.context uses absolute A1", async () => {
    const adapter = new WpsJsaAdapter();
    const defaultSpill = await adapter.readRange("Sheet1", "A1");
    expect(defaultSpill.ok).toBe(false);
    if (!defaultSpill.ok) {
      expect(defaultSpill.unsupported).toBe(true);
      expect(defaultSpill.reason).toContain("spill");
    }

    const expand = await adapter.readRange("Sheet1", "A1", "spill");
    expect(expand.ok).toBe(false);

    const none = await adapter.readRange("Sheet1", "B2:C3", "none");
    expect(none.ok).toBe(true);

    const ctx = await adapter.getFormulaContext("Sheet1", "B2:C3");
    expect(ctx.ok).toBe(true);
    if (ctx.ok) {
      expect(ctx.data.formulas.map((f) => f.address)).toEqual(
        expect.arrayContaining(["B2", "C3"]),
      );
      expect(ctx.data.formulas.every((f) => /^[A-Z]+\d+$/.test(f.address))).toBe(true);
    }

    expect((await adapter.copySheet("Sheet1")).ok).toBe(false);
    expect((await adapter.moveSheet("Sheet1", 1)).ok).toBe(false);
  });
});
