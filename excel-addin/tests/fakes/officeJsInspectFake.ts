/** Fake Excel.run for workbook.inspect per-sheet used-range dimensions. */
export function installInspectExcel() {
  type SheetState = {
    name: string;
    position: number;
    used: { address: string; rowCount: number; columnCount: number } | null;
  };

  const sheets = new Map<string, SheetState>();
  sheets.set("Sheet1", {
    name: "Sheet1",
    position: 0,
    used: { address: "Sheet1!A1:C3", rowCount: 3, columnCount: 3 },
  });
  sheets.set("Empty", {
    name: "Empty",
    position: 1,
    used: null,
  });
  sheets.set("Data", {
    name: "Data",
    position: 2,
    used: { address: "Data!B2:D10", rowCount: 9, columnCount: 3 },
  });

  let worksheetItems: ReturnType<typeof makeSheet>[] = [];
  let pendingItems: ReturnType<typeof makeSheet>[] | null = null;

  function makeUsed(sheet: SheetState) {
    if (!sheet.used) {
      return {
        isNullObject: true,
        address: "",
        rowCount: 0,
        columnCount: 0,
        load() {},
      };
    }
    return {
      isNullObject: false,
      address: sheet.used.address,
      rowCount: sheet.used.rowCount,
      columnCount: sheet.used.columnCount,
      load() {},
    };
  }

  function makeSheet(name: string) {
    const sheet = sheets.get(name);
    if (!sheet) throw new Error(`missing sheet ${name}`);
    return {
      get name() {
        return sheet.name;
      },
      get position() {
        return sheet.position;
      },
      load() {},
      getUsedRangeOrNullObject(_valuesOnly?: boolean) {
        return makeUsed(sheet);
      },
    };
  }

  const context = {
    workbook: {
      name: "Book1.xlsx",
      load() {},
      worksheets: {
        get items() {
          return worksheetItems;
        },
        load() {
          pendingItems = [...sheets.keys()].map((name) => makeSheet(name));
        },
        getActiveWorksheet() {
          return makeSheet("Sheet1");
        },
        getItem(name: string) {
          return makeSheet(name);
        },
      },
    },
    async sync() {
      if (pendingItems) {
        worksheetItems = pendingItems;
        pendingItems = null;
      }
    },
  };

  (globalThis as unknown as { window: unknown }).window = globalThis;
  (globalThis as unknown as { Excel: { run: Function } }).Excel = {
    run: async <T>(fn: (ctx: typeof context) => Promise<T>) => fn(context),
  };

  return { sheets };
}
