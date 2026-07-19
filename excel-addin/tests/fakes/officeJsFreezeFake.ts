import { toA1 } from "../../shared/host/a1Address";

/** Sync-gated fake: freeze mutations stay pending until context.sync(). */
export function installFreezeExcel() {
  type Freeze = { address: string; rowCount: number; columnCount: number } | null;
  type SheetState = {
    name: string;
    committed: Freeze;
    pending: Freeze | undefined;
  };

  const sheets = new Map<string, SheetState>();
  sheets.set("Sheet1", { name: "Sheet1", committed: null, pending: undefined });
  sheets.set("Sheet2", { name: "Sheet2", committed: null, pending: undefined });

  function makeLocation(sheet: SheetState) {
    // Properties resolve from committed state so load()+sync() can refresh them.
    return {
      load() {},
      get isNullObject() {
        return sheet.committed === null;
      },
      get address() {
        return sheet.committed?.address ?? "";
      },
      get rowCount() {
        return sheet.committed?.rowCount ?? 0;
      },
      get columnCount() {
        return sheet.committed?.columnCount ?? 0;
      },
    };
  }

  function makeSheet(name: string) {
    const sheet = sheets.get(name);
    if (!sheet) throw new Error(`missing sheet ${name}`);
    return {
      get name() {
        return sheet.name;
      },
      load() {},
      getRange(address: string) {
        return { address: `${name}!${address}`, load() {} };
      },
      freezePanes: {
        freezeRows(count: number) {
          sheet.pending = {
            address: `${name}!${toA1(count, 0)}`,
            rowCount: count,
            columnCount: 0,
          };
        },
        freezeColumns(count: number) {
          sheet.pending = {
            address: `${name}!${toA1(0, count - 1)}`,
            rowCount: 0,
            columnCount: count,
          };
        },
        freezeAt(range: { address?: string }) {
          const raw = String(range.address ?? "B2");
          const address = raw.includes("!") ? raw : `${name}!${raw}`;
          sheet.pending = { address, rowCount: 1, columnCount: 1 };
        },
        unfreeze() {
          sheet.pending = null;
        },
        getLocationOrNullObject() {
          return makeLocation(sheet);
        },
      },
    };
  }

  const context = {
    workbook: {
      worksheets: {
        getItem(name: string) {
          return makeSheet(name);
        },
      },
    },
    async sync() {
      for (const sheet of sheets.values()) {
        if (sheet.pending !== undefined) {
          sheet.committed = sheet.pending;
          sheet.pending = undefined;
        }
      }
    },
  };

  (globalThis as unknown as { window: unknown }).window = globalThis;
  (globalThis as unknown as { Excel: { run: Function } }).Excel = {
    run: async <T>(fn: (ctx: typeof context) => Promise<T>) => fn(context),
  };

  return {
    getCommitted(name: string) {
      return sheets.get(name)?.committed ?? null;
    },
    getPending(name: string) {
      return sheets.get(name)?.pending;
    },
  };
}
