/**
 * Sync-gated fake for Range.insert/delete and RangeFormat.autofit*.
 * Mutations commit on sync; loaded address/dimensions are only readable after
 * the follow-up load + sync, so tests catch input-echo and stale-read bugs.
 */

type Options = {
  excelApi12?: boolean;
  missingIsSetSupported?: boolean;
  isSetSupportedThrows?: boolean;
  missingInsert?: boolean;
  missingDelete?: boolean;
  missingAutofitColumns?: boolean;
  missingAutofitRows?: boolean;
  insertedAddress?: unknown;
  deletedAddress?: unknown;
  autofitAddress?: unknown;
  columnWidth?: unknown;
  rowHeight?: unknown;
  syncFailsAt?: number;
};

type Syncable = () => void;

export function installRangeStructureExcel(options: Options = {}) {
  let syncCount = 0;
  let runCalls = 0;
  let insertShift: string | null = null;
  let deleteShift: string | null = null;
  let autofitColumnsCalls = 0;
  let autofitRowsCalls = 0;
  const mutations: Syncable[] = [];
  const observers: Syncable[] = [];

  function makeRange(kind: "source" | "inserted", sourceAddress?: unknown) {
    let addressLoaded = false;
    let addressPending = false;
    let dimensionsLoaded = false;
    let dimensionsPending = false;
    const state = { mutationCommitted: kind === "source" };
    const requiresMutation = kind === "inserted";
    const address =
      kind === "inserted"
        ? (options.insertedAddress ?? "HostSheet!$A$1:$B$2")
        : (sourceAddress ?? options.autofitAddress ?? options.deletedAddress ?? "HostSheet!$C$3:$D$4");

    const range: {
      readonly address: unknown;
      load: (props?: string) => void;
      format: {
        readonly columnWidth: unknown;
        readonly rowHeight: unknown;
        load: (props?: string) => void;
        autofitColumns?: () => void;
        autofitRows?: () => void;
      };
      insert?: (shift: string) => unknown;
      delete?: (shift: string) => void;
    } = {
      get address() {
        if (!addressLoaded) throw new Error("Range.address not loaded");
        return address;
      },
      load(props = "") {
        if (props.includes("address")) addressPending = true;
      },
      format: {
        get columnWidth() {
          if (!dimensionsLoaded) throw new Error("columnWidth not loaded");
          return options.columnWidth ?? 73.5;
        },
        get rowHeight() {
          if (!dimensionsLoaded) throw new Error("rowHeight not loaded");
          return options.rowHeight ?? 21.25;
        },
        load(props = "") {
          if (props.includes("columnWidth") || props.includes("rowHeight")) {
            dimensionsPending = true;
          }
        },
      },
    };

    if (!options.missingInsert) {
      range.insert = (shift: string) => {
        insertShift = shift;
        const inserted = makeRange("inserted");
        mutations.push(() => {
          (inserted as { markMutationCommitted?: () => void }).markMutationCommitted?.();
        });
        return inserted;
      };
    }
    if (!options.missingDelete) {
      range.delete = (shift: string) => {
        deleteShift = shift;
        mutations.push(() => {
          state.mutationCommitted = true;
        });
      };
    }
    if (!options.missingAutofitColumns) {
      range.format.autofitColumns = () => {
        autofitColumnsCalls += 1;
        mutations.push(() => {
          state.mutationCommitted = true;
        });
      };
    }
    if (!options.missingAutofitRows) {
      range.format.autofitRows = () => {
        autofitRowsCalls += 1;
        mutations.push(() => {
          state.mutationCommitted = true;
        });
      };
    }

    const onSync = () => {
      if (addressPending && (!requiresMutation || state.mutationCommitted)) addressLoaded = true;
      if (dimensionsPending && state.mutationCommitted) dimensionsLoaded = true;
      addressPending = false;
      dimensionsPending = false;
    };
    observers.push(onSync);
    (range as { markMutationCommitted?: () => void }).markMutationCommitted = () => {
      state.mutationCommitted = true;
    };
    return range;
  }

  const context = {
    workbook: {
      worksheets: {
        getItem(name: string) {
          if (name !== "Sheet1") throw new Error(`missing sheet ${name}`);
          return {
            getRange(address: string) {
              if (!address) throw new Error("missing range");
              const hostAddress = /^C3(?::D4)?$/i.test(address)
                ? options.deletedAddress
                : options.autofitAddress;
              return makeRange("source", hostAddress);
            },
          };
        },
      },
    },
    async sync() {
      syncCount += 1;
      if (options.syncFailsAt === syncCount) throw new Error("sync failed");
      const callbacks = mutations.splice(0, mutations.length);
      for (const callback of callbacks) callback();
      for (const observer of observers) observer();
    },
  };

  (globalThis as unknown as { window: unknown }).window = globalThis;
  let requirements: Record<string, unknown> = {};
  if (!options.missingIsSetSupported) {
    requirements = {
      isSetSupported(name: string, version?: string) {
        if (options.isSetSupportedThrows) throw new Error("isSetSupported threw");
        return name === "ExcelApi" && version === "1.2" && options.excelApi12 !== false;
      },
    };
  }
  (globalThis as unknown as { Office: unknown }).Office = {
    context: { requirements },
  };
  (globalThis as unknown as { Excel: unknown }).Excel = {
    run: async <T>(fn: (value: typeof context) => Promise<T>) => {
      runCalls += 1;
      return fn(context);
    },
  };

  return {
    runCalls: () => runCalls,
    syncCalls: () => syncCount,
    insertShift: () => insertShift,
    deleteShift: () => deleteShift,
    autofitColumnsCalls: () => autofitColumnsCalls,
    autofitRowsCalls: () => autofitRowsCalls,
  };
}
