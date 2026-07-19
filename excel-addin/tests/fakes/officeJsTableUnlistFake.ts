/**
 * Sync-gated fake for Table.convertToRange().
 * Convert is pending until context.sync(); collection absence only after that sync.
 * Cell values stay in sheet.values after unlist.
 */

type TableState = {
  name: string;
  address: string;
  hostName?: string;
};

export function installTableUnlistExcel(options?: {
  excelApi12?: boolean;
  /** When true, isSetSupported throws. */
  isSetSupportedThrows?: boolean;
  /** When true, omit Office.context.requirements.isSetSupported. */
  missingIsSetSupported?: boolean;
  /** When false, table has no convertToRange. */
  hasConvertToRange?: boolean;
  hostSheetName?: string;
  tableName?: string;
  hostTableName?: string;
  address?: string;
  /** Cell values retained after convert (keyed by bare A1). */
  cellValues?: Record<string, unknown[][]>;
}) {
  const excelApi12 = options?.excelApi12 !== false;
  const hasConvertToRange = options?.hasConvertToRange !== false;
  let hostSheetName = options?.hostSheetName ?? "HostSheet";
  const lookupTableName = options?.tableName ?? "Sales";
  let hostTableName = options?.hostTableName ?? "HostSales";
  let hostAddress = options?.address ?? "HostSheet!A1:B3";
  const cellValues: Record<string, unknown[][]> = {
    A1: [["H1", "H2"], ["a", 1], ["b", 2]],
    ...(options?.cellValues ?? {}),
  };

  let tables = new Map<string, TableState>([
    [
      lookupTableName,
      {
        name: hostTableName,
        address: hostAddress,
      },
    ],
  ]);
  let pendingRemove: string | null = null;
  let convertCalls = 0;

  function makeTableApi(lookupName: string, state: TableState) {
    const api: {
      name: string;
      load: (_props?: string) => void;
      getRange: () => {
        address: string;
        load: (_props?: string) => void;
        values: unknown[][];
      };
      convertToRange?: () => { address: string };
      delete: () => void;
    } = {
      get name() {
        return state.name;
      },
      load(_props?: string) {},
      getRange() {
        return {
          get address() {
            return state.address;
          },
          load(_props?: string) {},
          get values() {
            const bare = state.address.includes("!")
              ? state.address.split("!")[1]!
              : state.address;
            return cellValues[bare] ?? cellValues.A1 ?? [];
          },
        };
      },
      delete() {
        tables.delete(lookupName);
        for (const [key, t] of tables) {
          if (t.name === state.name) tables.delete(key);
        }
      },
    };
    if (hasConvertToRange) {
      api.convertToRange = () => {
        convertCalls += 1;
        pendingRemove = lookupName;
        return { address: state.address };
      };
    }
    return api;
  }

  const context = {
    workbook: {
      worksheets: {
        getItem(name: string) {
          if (name !== "Sheet1" && name !== hostSheetName) {
            // allow request sheet key Sheet1
            if (name !== "Sheet1") throw new Error(`missing sheet ${name}`);
          }
          return {
            get name() {
              return hostSheetName;
            },
            load(_props?: string) {},
            tables: {
              get items() {
                return [...tables.values()].map((t) => ({
                  name: t.name,
                  load(_props?: string) {},
                }));
              },
              load(_props?: string) {},
              getItem(tableName: string) {
                const state = tables.get(tableName);
                if (!state) {
                  for (const [key, t] of tables) {
                    if (t.name.toLowerCase() === tableName.toLowerCase()) {
                      return makeTableApi(key, t);
                    }
                  }
                  throw new Error(`missing table ${tableName}`);
                }
                return makeTableApi(tableName, state);
              },
            },
            getRange(address: string) {
              const bare = address.includes("!") ? address.split("!")[1]! : address;
              return {
                address: `${hostSheetName}!${bare}`,
                values: cellValues[bare] ?? cellValues.A1 ?? [],
                load(_props?: string) {},
              };
            },
          };
        },
      },
    },
    async sync() {
      if (pendingRemove != null) {
        tables.delete(pendingRemove);
        // also drop by host name match
        for (const [key, t] of [...tables.entries()]) {
          if (t.name === hostTableName) tables.delete(key);
        }
        pendingRemove = null;
      }
    },
  };

  (globalThis as unknown as { window: unknown }).window = globalThis;

  if (options?.missingIsSetSupported) {
    (globalThis as unknown as {
      Office: { context: { requirements: Record<string, unknown> } };
    }).Office = {
      context: { requirements: {} },
    };
  } else if (options?.isSetSupportedThrows) {
    (globalThis as unknown as {
      Office: {
        context: {
          requirements: { isSetSupported: () => boolean };
        };
      };
    }).Office = {
      context: {
        requirements: {
          isSetSupported() {
            throw new Error("isSetSupported threw");
          },
        },
      },
    };
  } else {
    (globalThis as unknown as {
      Office: {
        context: {
          requirements: { isSetSupported: (name: string, minVersion?: string) => boolean };
        };
      };
    }).Office = {
      context: {
        requirements: {
          isSetSupported(name: string, minVersion?: string) {
            if (name === "ExcelApi" && minVersion === "1.2") return excelApi12;
            return false;
          },
        },
      },
    };
  }

  (globalThis as unknown as { Excel: { run: Function } }).Excel = {
    run: async <T>(fn: (ctx: typeof context) => Promise<T>) => fn(context),
  };

  return {
    convertCalls() {
      return convertCalls;
    },
    tableNames() {
      return [...tables.values()].map((t) => t.name);
    },
    cellValuesAt(address: string) {
      const bare = address.includes("!") ? address.split("!")[1]! : address;
      return cellValues[bare] ?? cellValues.A1;
    },
    /** Skip convert-commit sync: convert pending, collection still has old table. */
    async brokenSkipConvertSync() {
      const sheet = context.workbook.worksheets.getItem("Sheet1");
      const table = sheet.tables.getItem(lookupTableName);
      sheet.load("name");
      table.load("name");
      const range = table.getRange();
      range.load("address");
      await context.sync();
      if (typeof table.convertToRange !== "function") {
        throw new Error("no convertToRange");
      }
      table.convertToRange();
      // intentionally skip sync that would commit remove
      sheet.tables.load("items/name");
      const names = sheet.tables.items.map((t) => t.name);
      // abandon uncommitted pending so later full runs stay independent
      pendingRemove = null;
      return names;
    },
    setHostNames(sheet: string, table: string, address: string) {
      hostSheetName = sheet;
      hostTableName = table;
      hostAddress = address;
      tables = new Map([
        [lookupTableName, { name: hostTableName, address: hostAddress }],
      ]);
    },
  };
}
