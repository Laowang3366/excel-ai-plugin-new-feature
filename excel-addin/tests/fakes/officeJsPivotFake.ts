/**
 * Sync-gated Office.js PivotTable fake (ExcelApi 1.8 surface).
 */
import {
  createLoadTracker,
  installPivotOfficeRequirements,
  makeRange,
  type HierarchyAxis,
  type InstallPivotExcelOptions,
  type Pending,
  type PivotState,
  type SheetState,
} from "./officeJsPivotFakeHelpers";

export type { InstallPivotExcelOptions };

export function installPivotExcel(options: InstallPivotExcelOptions = {}) {
  const excelApi18 = options.excelApi18 !== false;
  const excelApi13 = options.excelApi13 !== false;
  const excelApi17 = options.excelApi17 !== false;
  const hierarchyNames = options.hierarchyNames ?? ["Region", "Product", "Sales"];
  const sheetNames = options.sheets ?? ["Sheet1"];
  const strictLoad = options.strictLoad !== false;
  const { markLoaded, requireLoaded: requireLoadedRaw } = createLoadTracker();
  const requireLoaded = (obj: object, prop: string) => requireLoadedRaw(obj, prop, strictLoad);

  const sheets = new Map<string, SheetState>();
  for (const name of sheetNames) {
    sheets.set(name, { name, pivots: new Map() });
  }

  let syncCount = 0;
  let addCalls = 0;
  let refreshCalls = 0;
  let connectionRefreshCalls = 0;
  let pending: Pending[] = [];

  function makeHierarchyCollection(pivot: PivotState, axis: HierarchyAxis) {
    const coll = {
      items: [] as Array<{ name: string; summarizeBy?: string }>,
      load(props?: string) {
        if (props) markLoaded(coll, props);
        markLoaded(coll, "items");
        refreshItems();
      },
      add(hierarchy: { name: string }) {
        const field = hierarchy.name;
        if (axis === "data") {
          const dataEntry = { name: field, summarizeBy: "Sum", caption: field };
          pending.push({ kind: "layout", pivot, axis, field, data: dataEntry });
          const dataProxy: {
            name: string;
            summarizeBy: string;
            load: (props?: string) => void;
          } = {
            name: field,
            summarizeBy: "Sum",
            load(props?: string) {
              if (props) markLoaded(dataProxy, props);
            },
          };
          Object.defineProperty(dataProxy, "name", {
            get() {
              return dataEntry.caption;
            },
            set(v: string) {
              dataEntry.caption = v;
            },
            configurable: true,
          });
          Object.defineProperty(dataProxy, "summarizeBy", {
            get() {
              return dataEntry.summarizeBy;
            },
            set(v: string) {
              dataEntry.summarizeBy = v;
            },
            configurable: true,
          });
          return dataProxy;
        }
        pending.push({ kind: "layout", pivot, axis, field });
        return hierarchy;
      },
      getItem(name: string) {
        return { name, load(_p?: string) {} };
      },
    };

    function refreshItems() {
      if (axis === "row") coll.items = pivot.row.map((n) => ({ name: n }));
      else if (axis === "column") coll.items = pivot.column.map((n) => ({ name: n }));
      else if (axis === "filter") coll.items = pivot.filter.map((n) => ({ name: n }));
      else {
        coll.items = pivot.data.map((d) => ({
          name: d.caption || d.name,
          summarizeBy: d.summarizeBy,
        }));
      }
    }
    return coll;
  }

  function makePivot(state: PivotState) {
    const pivot: Record<string, unknown> = {};
    const layoutRange = makeRange(state.destAddress, state.sheetName, requireLoaded, markLoaded);
    layoutRange.rowCount = Math.max(5, state.layoutEndRow);

    Object.defineProperty(pivot, "name", {
      get() {
        requireLoaded(pivot, "name");
        return state.name;
      },
      configurable: true,
    });

    const worksheet = {
      get name() {
        requireLoaded(worksheet, "name");
        return state.sheetName;
      },
      load(props?: string) {
        if (props) markLoaded(worksheet, props);
      },
    };

    pivot.worksheet = worksheet;
    pivot.hierarchies = {
      getItem(name: string) {
        return { name, load(_p?: string) {} };
      },
      load(_p?: string) {},
      items: hierarchyNames.map((n) => ({ name: n })),
    };
    pivot.rowHierarchies = makeHierarchyCollection(state, "row");
    pivot.columnHierarchies = makeHierarchyCollection(state, "column");
    pivot.filterHierarchies = makeHierarchyCollection(state, "filter");
    pivot.dataHierarchies = makeHierarchyCollection(state, "data");
    pivot.layout = { getRange: () => layoutRange };
    pivot.load = (props?: string) => {
      if (props) markLoaded(pivot, props);
    };
    if (!options.noDataSourceString) {
      pivot.getDataSourceString = () => state.sourceAddress;
    }
    if (!options.missingRefresh) {
      pivot.refresh = () => {
        pending.push({ kind: "refresh", pivot: state });
      };
    }
    return pivot;
  }

  function makeSheet(state: SheetState) {
    const sheet: Record<string, unknown> = {
      get name() {
        requireLoaded(sheet, "name");
        return state.name;
      },
      load(props?: string) {
        if (props) markLoaded(sheet, props);
      },
      getRange(address: string) {
        return makeRange(address, state.name, requireLoaded, markLoaded);
      },
      pivotTables: {
        items: [] as unknown[],
        load(props?: string) {
          if (props) markLoaded(this as object, props);
          (this as { items: unknown[] }).items = [...state.pivots.values()].map((p) => {
            const pivotObj = makePivot(p);
            if (props && (props.includes("name") || props.includes("items"))) {
              (pivotObj as { load: (p?: string) => void }).load("name");
            }
            return pivotObj;
          });
        },
        getItem(name: string) {
          const found = state.pivots.get(name);
          if (!found) {
            for (const [k, v] of state.pivots) {
              if (k.toLowerCase() === name.toLowerCase()) return makePivot(v);
            }
            throw new Error(`pivot not found: ${name}`);
          }
          return makePivot(found);
        },
        add: options.missingAdd
          ? undefined
          : (name: string, source: { address?: string }, destination: { address?: string }) => {
              addCalls += 1;
              const srcBare = "A1:C10";
              const destBare =
                typeof destination === "object" && destination
                  ? String((destination as { address?: string }).address ?? "A1")
                      .split("!")
                      .pop()!
                  : "A1";
              const pivotState: PivotState = {
                name,
                sheetName: state.name,
                sourceAddress: `${sheetNames[0]}!${srcBare}`,
                destAddress: destBare.replace(/\$/g, "").toUpperCase().split(":")[0] ?? "A1",
                hierarchyNames: [...hierarchyNames],
                row: [],
                column: [],
                filter: [],
                data: [],
                refreshed: false,
                layoutEndRow: 8,
              };
              try {
                const a = (source as { address: string }).address;
                if (a) pivotState.sourceAddress = a;
              } catch {
                /* not loaded yet */
              }
              pending.push({ kind: "add", sheet: state.name, pivot: pivotState });
              return makePivot(pivotState);
            },
        refreshAll() {
          for (const p of state.pivots.values()) {
            pending.push({ kind: "refresh", pivot: p });
          }
        },
      },
    };
    return sheet;
  }

  async function sync() {
    syncCount += 1;
    const batch = pending;
    pending = [];
    for (const op of batch) {
      if (op.kind === "createSheet") {
        if (!sheets.has(op.name)) sheets.set(op.name, { name: op.name, pivots: new Map() });
      } else if (op.kind === "add") {
        const sheet = sheets.get(op.sheet);
        if (!sheet) throw new Error(`sheet missing: ${op.sheet}`);
        if (options.tamperHierarchies) {
          op.pivot.row = [];
          op.pivot.column = [];
          op.pivot.filter = [];
          op.pivot.data = [];
        }
        sheet.pivots.set(op.pivot.name, op.pivot);
      } else if (op.kind === "layout") {
        if (options.tamperHierarchies) continue;
        if (op.axis === "row" && !op.pivot.row.includes(op.field)) op.pivot.row.push(op.field);
        if (op.axis === "column" && !op.pivot.column.includes(op.field)) {
          op.pivot.column.push(op.field);
        }
        if (op.axis === "filter" && !op.pivot.filter.includes(op.field)) {
          op.pivot.filter.push(op.field);
        }
        if (op.axis === "data" && op.data) {
          // Desktop allows multiple data aggregates on the same source field.
          const d = op.data as { name: string; summarizeBy: string; caption: string };
          op.pivot.data.push({
            name: d.name,
            summarizeBy: d.summarizeBy,
            caption: d.caption,
          });
        }
      } else if (op.kind === "refresh") {
        refreshCalls += 1;
        op.pivot.refreshed = true;
      } else if (op.kind === "connectionRefreshAll") {
        connectionRefreshCalls += 1;
      }
    }
  }

  const dataConnections = options.missingDataConnections
    ? undefined
    : {
        refreshAll: options.missingRefreshAll
          ? undefined
          : () => {
              pending.push({ kind: "connectionRefreshAll" });
            },
      };

  const context = {
    workbook: {
      name: "Book1.xlsx",
      load(_p?: string) {},
      ...(dataConnections ? { dataConnections } : {}),
      worksheets: {
        items: [] as unknown[],
        load(props?: string) {
          if (props) markLoaded(this, props);
          this.items = [...sheets.values()].map((s) => {
            const sheetObj = makeSheet(s);
            if (props && props.includes("name")) {
              (sheetObj as { load: (p?: string) => void }).load("name");
            }
            return sheetObj;
          });
        },
        getItem(name: string) {
          const found = sheets.get(name);
          if (!found) {
            for (const [k, v] of sheets) {
              if (k.toLowerCase() === name.toLowerCase()) return makeSheet(v);
            }
            throw new Error(`sheet not found: ${name}`);
          }
          return makeSheet(found);
        },
        add(name?: string) {
          const n = name ?? `Sheet${sheets.size + 1}`;
          pending.push({ kind: "createSheet", name: n });
          if (!sheets.has(n)) sheets.set(n, { name: n, pivots: new Map() });
          return makeSheet(sheets.get(n)!);
        },
        getActiveWorksheet() {
          return makeSheet([...sheets.values()][0]!);
        },
      },
    },
    sync,
  };

  const g = globalThis as unknown as {
    Excel?: { run: <T>(fn: (ctx: typeof context) => Promise<T>) => Promise<T> };
  };
  g.Excel = { run: async (fn) => fn(context) };
  installPivotOfficeRequirements({
    excelApi18,
    excelApi17,
    excelApi13,
    missingIsSetSupported: options.missingIsSetSupported,
    isSetSupportedThrows: options.isSetSupportedThrows,
  });

  return {
    syncCount: () => syncCount,
    addCalls: () => addCalls,
    refreshCalls: () => refreshCalls,
    connectionRefreshCalls: () => connectionRefreshCalls,
    pivotNames: (sheet?: string) => {
      const out: string[] = [];
      for (const s of sheets.values()) {
        if (sheet && s.name !== sheet) continue;
        out.push(...s.pivots.keys());
      }
      return out;
    },
    getPivot: (name: string) => {
      for (const s of sheets.values()) {
        const p = s.pivots.get(name);
        if (p) return { ...p, data: [...p.data], row: [...p.row] };
      }
      return null;
    },
    seedPivot(state: Partial<PivotState> & { name: string; sheetName?: string }) {
      const sheetName = state.sheetName ?? "Sheet1";
      if (!sheets.has(sheetName)) sheets.set(sheetName, { name: sheetName, pivots: new Map() });
      const full: PivotState = {
        name: state.name,
        sheetName,
        sourceAddress: state.sourceAddress ?? "Sheet1!A1:C10",
        destAddress: state.destAddress ?? "A1",
        hierarchyNames: state.hierarchyNames ?? hierarchyNames,
        row: state.row ?? [],
        column: state.column ?? [],
        filter: state.filter ?? [],
        data: state.data ?? [],
        refreshed: state.refreshed ?? false,
        layoutEndRow: state.layoutEndRow ?? 8,
      };
      sheets.get(sheetName)!.pivots.set(state.name, full);
    },
  };
}
