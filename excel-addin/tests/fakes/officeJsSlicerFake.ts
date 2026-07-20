/**
 * Install minimal Excel.run + Workbook/Worksheet.slicers fake for Phase54 tests.
 */
import {
  commitLoads,
  createLoadMaps,
  defaultItems,
  flushClientResults,
  makeSlicerProxy,
  markLoad,
  requireLoaded,
  type ClientResultBag,
  type SlicerFakeState,
  type WriteCounts,
} from "./officeJsSlicerFakeHelpers";

export type SlicerFakeInstall = {
  uninstall: () => void;
  state: {
    slicers: SlicerFakeState[];
    writeCounts: WriteCounts;
    syncCount: number;
    excelRunCalls: number;
    setRequirement: (supported: boolean | "throw" | "missing") => void;
    poisonSortBy: (name: string, value: string) => void;
    poisonScalar: (name: string, field: string, value: unknown) => void;
    poisonItem: (name: string, index: number, field: string, value: unknown) => void;
    removeGetSelectedItems: (name: string) => void;
    setSelectItemsNoOp: (name: string) => void;
    disableGetItemOrNullObject: () => void;
    lastClientResultReadBeforeSync: boolean;
  };
};

export function installSlicerExcel(options?: {
  excelApi110?: boolean | "throw" | "missing";
}): SlicerFakeInstall {
  const previous = {
    Excel: (globalThis as { Excel?: unknown }).Excel,
    Office: (globalThis as { Office?: unknown }).Office,
    window: (globalThis as { window?: unknown }).window,
  };

  const maps = createLoadMaps();
  const writeCounts: WriteCounts = {
    add: 0,
    delete: 0,
    selectItems: 0,
    clearFilters: 0,
    propertySets: 0,
  };
  const slicers: SlicerFakeState[] = [];
  const tables = [{ name: "SalesTable" }];
  const pivots = [{ name: "SalesPivot" }];
  const sheets = ["Sheet1", "Sheet2"];
  let syncCount = 0;
  let excelRunCalls = 0;
  let requirement: boolean | "throw" | "missing" = options?.excelApi110 ?? true;
  const poisonSortBy = new Map<string, string>();
  const noGetSelected = new Set<string>();
  const selectItemsNoOp = new Set<string>();
  let useGetItemOrNullObject = true;
  let lastClientResultReadBeforeSync = false;
  const clientResults: ClientResultBag = { pending: [], generation: 0 };

  function proxyOpts(s: SlicerFakeState) {
    return {
      poisonSortBy,
      noGetSelected,
      selectItemsNoOp,
      clientResults,
      onDelete: () => {
        const idx = slicers.indexOf(s);
        if (idx >= 0) slicers.splice(idx, 1);
      },
    };
  }

  function makeCollection(filterSheet?: string) {
    const coll: {
      items: object[];
      load(props: string): void;
      add(source: unknown, field: unknown, dest?: unknown): object;
      getItem(key: string): object;
      getItemOrNullObject?: (key: string) => object;
    } = {
      items: [],
      load(props: string) {
        markLoad(maps, coll, props);
      },
      add(_source: unknown, _field: unknown, dest?: unknown) {
        writeCounts.add += 1;
        let sheetName = "Sheet1";
        if (dest && typeof dest === "object" && dest !== null) {
          try {
            const n = (dest as { name?: string }).name;
            if (typeof n === "string") sheetName = n;
          } catch {
            sheetName = "Sheet1";
          }
        } else if (typeof dest === "string") sheetName = dest;
        const s: SlicerFakeState = {
          id: `id-${slicers.length + 1}`,
          name: `Slicer${slicers.length + 1}`,
          caption: "",
          sheetName,
          top: 0,
          left: 0,
          width: 100,
          height: 150,
          sortBy: "DataSourceOrder",
          style: "SlicerStyleLight1",
          isFilterCleared: true,
          items: defaultItems(),
        };
        slicers.push(s);
        return makeSlicerProxy(s, maps, writeCounts, proxyOpts(s));
      },
      getItem(key: string) {
        const found = slicers.find(
          (x) =>
            (x.name === key || x.id === key) &&
            (filterSheet === undefined || x.sheetName === filterSheet),
        );
        if (!found) throw new Error(`slicer not found: ${key}`);
        return makeSlicerProxy(found, maps, writeCounts, proxyOpts(found));
      },
    };
    maps.proxies.push(coll);
    Object.defineProperty(coll, "items", {
      get() {
        requireLoaded(maps, coll, "items");
        return slicers
          .filter((s) => filterSheet === undefined || s.sheetName === filterSheet)
          .map((s) => makeSlicerProxy(s, maps, writeCounts, proxyOpts(s)));
      },
      configurable: true,
    });
    Object.defineProperty(coll, "getItemOrNullObject", {
      get() {
        if (!useGetItemOrNullObject) return undefined;
        return (key: string) => {
          const found = slicers.find(
            (x) =>
              (x.name === key || x.id === key) &&
              (filterSheet === undefined || x.sheetName === filterSheet),
          );
          if (!found) return { isNullObject: true, load() {}, name: "" };
          const proxy = makeSlicerProxy(found, maps, writeCounts, proxyOpts(found)) as Record<
            string,
            unknown
          >;
          proxy.isNullObject = false;
          return proxy;
        };
      },
      configurable: true,
      enumerable: true,
    });
    return coll;
  }

  function makeWorksheet(name: string) {
    const sheet: { load(props: string): void; name?: string; slicers?: object } = {
      load(props: string) {
        markLoad(maps, sheet, props);
      },
    };
    maps.proxies.push(sheet);
    Object.defineProperty(sheet, "name", {
      get() {
        requireLoaded(maps, sheet, "name");
        return name;
      },
      configurable: true,
    });
    Object.defineProperty(sheet, "slicers", {
      get() {
        return makeCollection(name);
      },
      configurable: true,
    });
    return sheet;
  }

  const workbookSlicers = makeCollection();

  const Excel = {
    run: async <T>(batch: (ctx: unknown) => Promise<T>): Promise<T> => {
      excelRunCalls += 1;
      const context = {
        sync: async () => {
          syncCount += 1;
          commitLoads(maps);
          flushClientResults(clientResults);
        },
        workbook: {
          slicers: workbookSlicers,
          tables: {
            getItem(name: string) {
              const t = tables.find((x) => x.name === name);
              if (!t) throw new Error(`table not found: ${name}`);
              const proxy = {
                name: t.name,
                load(props: string) {
                  markLoad(maps, proxy, props);
                },
              };
              maps.proxies.push(proxy);
              Object.defineProperty(proxy, "name", {
                get() {
                  requireLoaded(maps, proxy, "name");
                  return t.name;
                },
                configurable: true,
              });
              return proxy;
            },
          },
          pivotTables: {
            getItem(name: string) {
              const p = pivots.find((x) => x.name === name);
              if (!p) throw new Error(`pivot not found: ${name}`);
              const proxy = {
                name: p.name,
                load(props: string) {
                  markLoad(maps, proxy, props);
                },
              };
              maps.proxies.push(proxy);
              Object.defineProperty(proxy, "name", {
                get() {
                  requireLoaded(maps, proxy, "name");
                  return p.name;
                },
                configurable: true,
              });
              return proxy;
            },
          },
          worksheets: {
            getItem(name: string) {
              if (!sheets.includes(name)) throw new Error(`sheet not found: ${name}`);
              return makeWorksheet(name);
            },
            load(_props: string) {},
          },
        },
      };
      return batch(context);
    },
  };

  function bindIsSetSupported(): void {
    const req = (globalThis as {
      Office?: { context?: { requirements?: { isSetSupported?: unknown } } };
    }).Office?.context?.requirements;
    if (!req) return;
    if (requirement === "missing") {
      delete req.isSetSupported;
      return;
    }
    req.isSetSupported = (name: string, version?: string) => {
      if (requirement === "throw") throw new Error("isSetSupported threw");
      if (name === "ExcelApi" && version === "1.10") return requirement === true;
      return false;
    };
  }

  const Office = {
    context: {
      requirements: {
        isSetSupported(_name: string, _version?: string) {
          return false;
        },
      },
    },
  };

  (globalThis as { Excel?: unknown }).Excel = Excel;
  (globalThis as { Office?: unknown }).Office = Office;
  (globalThis as { window?: unknown }).window = globalThis;
  if (options?.excelApi110 === "missing") requirement = "missing";
  bindIsSetSupported();

  return {
    uninstall() {
      if (previous.Excel === undefined) delete (globalThis as { Excel?: unknown }).Excel;
      else (globalThis as { Excel?: unknown }).Excel = previous.Excel;
      if (previous.Office === undefined) delete (globalThis as { Office?: unknown }).Office;
      else (globalThis as { Office?: unknown }).Office = previous.Office;
      if (previous.window === undefined) delete (globalThis as { window?: unknown }).window;
      else (globalThis as { window?: unknown }).window = previous.window;
    },
    state: {
      slicers,
      writeCounts,
      get syncCount() {
        return syncCount;
      },
      get excelRunCalls() {
        return excelRunCalls;
      },
      get lastClientResultReadBeforeSync() {
        return lastClientResultReadBeforeSync;
      },
      setRequirement(supported: boolean | "throw" | "missing") {
        requirement = supported;
        bindIsSetSupported();
      },
      poisonSortBy(name: string, value: string) {
        poisonSortBy.set(name, value);
      },
      poisonScalar(name: string, field: string, value: unknown) {
        const s = slicers.find((x) => x.name === name);
        if (!s) throw new Error(`poisonScalar: slicer not found ${name}`);
        s.poison = { ...(s.poison ?? {}), [field]: value };
      },
      poisonItem(name: string, index: number, field: string, value: unknown) {
        const s = slicers.find((x) => x.name === name);
        if (!s) throw new Error(`poisonItem: slicer not found ${name}`);
        if (!s.itemPoison) s.itemPoison = [];
        s.itemPoison[index] = { ...(s.itemPoison[index] ?? {}), [field]: value };
      },
      removeGetSelectedItems(name: string) {
        noGetSelected.add(name);
      },
      setSelectItemsNoOp(name: string) {
        selectItemsNoOp.add(name);
      },
      disableGetItemOrNullObject() {
        useGetItemOrNullObject = false;
      },
    },
  };
}
