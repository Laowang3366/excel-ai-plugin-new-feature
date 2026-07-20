/** @MOCK_INTERFACE — Excel.run Table.sort double with apply/clear/fields load fence. */

export type TableSortFakeOptions = {
  excelApi12?: boolean;
  missingIsSetSupported?: boolean;
  isSetSupportedThrows?: boolean;
  hasApply?: boolean;
  hasClear?: boolean;
  hostSheetName?: string;
  hostTableName?: string;
};

type FieldState = { key: number; ascending: boolean };

export function installTableSortExcel(options: TableSortFakeOptions = {}) {
  const excelApi12 = options.excelApi12 !== false;
  const hasApply = options.hasApply !== false;
  const hasClear = options.hasClear !== false;
  const hostSheetName = options.hostSheetName ?? "HostSheet";
  const hostTableName = options.hostTableName ?? "HostSales";

  let applyCalls = 0;
  let clearCalls = 0;
  let fieldsCommitted: FieldState[] = [];
  let fieldsPending: FieldState[] | undefined;

  let sheetNameSnap: string | undefined;
  let sheetNamePending: string | undefined;
  let tableNameSnap: string | undefined;
  let tableNamePending: string | undefined;

  type FieldProxy = {
    key?: number;
    ascending?: boolean;
    load: (p: string) => void;
    _flushLoad: () => void;
  };

  let fieldItemsSnap: FieldProxy[] | undefined;
  let fieldItemsPending: FieldProxy[] | undefined;

  function makeFieldProxy(state: FieldState): FieldProxy {
    let pending: FieldState | undefined;
    let snap: FieldState | undefined;
    return {
      get key() {
        if (!snap) throw new Error("SortField.key not loaded");
        return snap.key;
      },
      get ascending() {
        if (!snap) throw new Error("SortField.ascending not loaded");
        return snap.ascending;
      },
      load() {
        pending = { ...state };
        snap = undefined;
      },
      _flushLoad() {
        if (pending) {
          snap = pending;
          pending = undefined;
        }
      },
    };
  }

  const fieldsApi = {
    items: [] as FieldProxy[],
    load(_props: string) {
      fieldItemsPending = fieldsCommitted.map((f) => makeFieldProxy(f));
      fieldItemsSnap = undefined;
      this.items = [];
    },
    _flushLoad() {
      if (fieldItemsPending) {
        fieldItemsSnap = fieldItemsPending;
        this.items = fieldItemsSnap;
        fieldItemsPending = undefined;
      }
      for (const item of this.items) item._flushLoad();
    },
  };

  const sortApi = {
    fields: fieldsApi,
    apply(fields: Array<{ key: number; ascending: boolean }>, _matchCase?: boolean) {
      if (!hasApply) throw new Error("Table.sort.apply missing");
      applyCalls += 1;
      fieldsPending = fields.map((f) => ({ key: f.key, ascending: f.ascending !== false }));
    },
    clear() {
      if (!hasClear) throw new Error("Table.sort.clear missing");
      clearCalls += 1;
      fieldsPending = [];
    },
  };

  const table = {
    get name() {
      if (tableNameSnap === undefined) throw new Error("Table.name not loaded");
      return tableNameSnap;
    },
    load(_props: string) {
      tableNamePending = hostTableName;
      tableNameSnap = undefined;
    },
    sort: sortApi,
    _flushLoad() {
      if (tableNamePending !== undefined) {
        tableNameSnap = tableNamePending;
        tableNamePending = undefined;
      }
      fieldsApi._flushLoad();
    },
  };

  const sheet = {
    get name() {
      if (sheetNameSnap === undefined) throw new Error("Worksheet.name not loaded");
      return sheetNameSnap;
    },
    load(_props: string) {
      sheetNamePending = hostSheetName;
      sheetNameSnap = undefined;
    },
    tables: {
      getItem(_name: string) {
        return table;
      },
    },
    _flushLoad() {
      if (sheetNamePending !== undefined) {
        sheetNameSnap = sheetNamePending;
        sheetNamePending = undefined;
      }
      table._flushLoad();
    },
  };

  async function sync() {
    if (fieldsPending !== undefined) {
      fieldsCommitted = fieldsPending;
      fieldsPending = undefined;
    }
    sheet._flushLoad();
  }

  (globalThis as { Excel?: unknown }).Excel = {
    run: async (cb: (context: unknown) => Promise<unknown>) =>
      cb({
        workbook: {
          worksheets: {
            getItem(_name: string) {
              return sheet;
            },
          },
        },
        sync,
      }),
  };

  (globalThis as unknown as { window: unknown }).window = globalThis;

  if (options.missingIsSetSupported) {
    (globalThis as { Office?: unknown }).Office = { context: { requirements: {} } };
  } else if (options.isSetSupportedThrows) {
    (globalThis as { Office?: unknown }).Office = {
      context: {
        requirements: {
          isSetSupported() {
            throw new Error("boom");
          },
        },
      },
    };
  } else {
    (globalThis as { Office?: unknown }).Office = {
      context: {
        requirements: {
          isSetSupported(_name: string, version?: string) {
            if (version === "1.2") return excelApi12;
            return false;
          },
        },
      },
    };
  }

  return {
    applyCalls: () => applyCalls,
    clearCalls: () => clearCalls,
    fields: () => fieldsCommitted.map((f) => ({ ...f })),
  };
}
