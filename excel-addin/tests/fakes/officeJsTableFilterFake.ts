/** @MOCK_INTERFACE — Excel.run Table.autoFilter double with apply/clear/enabled load fence. */

export type TableFilterFakeOptions = {
  excelApi12?: boolean;
  excelApi19?: boolean;
  missingIsSetSupported?: boolean;
  isSetSupportedThrows?: boolean;
  hasApply?: boolean;
  hasClear?: boolean;
  hostSheetName?: string;
  hostTableName?: string;
};

export function installTableFilterExcel(options: TableFilterFakeOptions = {}) {
  const excelApi12 = options.excelApi12 !== false;
  const excelApi19 = options.excelApi19 !== false;
  const hasApply = options.hasApply !== false;
  const hasClear = options.hasClear !== false;
  const hostSheetName = options.hostSheetName ?? "HostSheet";
  const hostTableName = options.hostTableName ?? "HostSales";

  let applyCalls = 0;
  let clearCalls = 0;
  let lastApply: { columnIndex: number; criteria: Record<string, unknown> } | null = null;
  let enabledCommitted = false;
  let enabledPending: boolean | undefined;

  let sheetNameSnap: string | undefined;
  let sheetNamePending: string | undefined;
  let tableNameSnap: string | undefined;
  let tableNamePending: string | undefined;
  let enabledSnap: boolean | undefined;
  let enabledLoadPending: boolean | undefined;

  const autoFilter: Record<string, unknown> = {
    get enabled() {
      if (enabledSnap === undefined) throw new Error("AutoFilter.enabled not loaded");
      return enabledSnap;
    },
    apply(_range: unknown, columnIndex: number, criteria: Record<string, unknown>) {
      if (!hasApply) throw new Error("Table.autoFilter.apply missing");
      applyCalls += 1;
      lastApply = { columnIndex, criteria };
      enabledPending = true;
      enabledCommitted = true;
    },
    clearCriteria() {
      if (!hasClear) throw new Error("Table.autoFilter.clearCriteria missing");
      clearCalls += 1;
      enabledPending = false;
      enabledCommitted = false;
    },
    load(_props: string) {
      enabledLoadPending = enabledCommitted;
      enabledSnap = undefined;
    },
    _flushLoad() {
      if (enabledLoadPending !== undefined) {
        enabledSnap = enabledLoadPending;
        enabledLoadPending = undefined;
      }
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
    getRange() {
      return { address: `${hostSheetName}!A1:C10` };
    },
    autoFilter,
    _flushLoad() {
      if (tableNamePending !== undefined) {
        tableNameSnap = tableNamePending;
        tableNamePending = undefined;
      }
      (autoFilter as { _flushLoad: () => void })._flushLoad();
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
    if (enabledPending !== undefined) {
      enabledCommitted = enabledPending;
      enabledPending = undefined;
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
            if (version === "1.9") return excelApi19;
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
    lastApply: () => lastApply,
    enabled: () => enabledCommitted,
  };
}
