/**
 * Sync-gated fake for Chart.setData + series readback.
 * setData writes stay pending until context.sync(); series.items only after load+sync.
 * Multi-sheet worksheets supported for cross-sheet source Range proxies.
 */
type SeriesState = { name: string; chartType: string; smooth: boolean };

type ChartState = {
  name: string;
  pending?: { sourceRange: string; seriesBy: string; series: SeriesState[]; sourceSheet: string };
  committed: { sourceRange: string; seriesBy: string; series: SeriesState[]; sourceSheet: string };
};

type SeriesProxy = {
  name: string;
  chartType: string;
  smooth: boolean;
};

export function installChartSourceExcel(options?: {
  sheets?: string[];
  chartSheet?: string;
  chartName?: string;
  /** When false, chart has no setData method. */
  withSetData?: boolean;
}) {
  const sheetNames = options?.sheets ?? ["Sheet1", "Sheet2", "Sheet 2"];
  const chartSheetName = options?.chartSheet ?? "Sheet1";
  const chartName = options?.chartName ?? "C1";
  const withSetData = options?.withSetData !== false;

  const chart: ChartState = {
    name: chartName,
    committed: {
      sourceRange: "A1:B2",
      seriesBy: "Auto",
      sourceSheet: chartSheetName,
      series: [
        { name: "Series1", chartType: "ColumnClustered", smooth: false },
        { name: "Series2", chartType: "ColumnClustered", smooth: false },
      ],
    },
  };

  let seriesItems: SeriesProxy[] | null = null;
  let pendingSeriesItems: SeriesProxy[] | null = null;
  const seriesSnapshots = new WeakMap<SeriesProxy, SeriesState>();

  function makeSeriesProxy(snap: SeriesState): SeriesProxy {
    const proxy = {} as SeriesProxy;
    Object.defineProperties(proxy, {
      name: {
        get() {
          return seriesSnapshots.get(proxy)?.name ?? snap.name;
        },
      },
      chartType: {
        get() {
          return seriesSnapshots.get(proxy)?.chartType ?? snap.chartType;
        },
      },
      smooth: {
        get() {
          return seriesSnapshots.get(proxy)?.smooth ?? snap.smooth;
        },
      },
    });
    return proxy;
  }

  function buildSeriesFromSource(sourceRange: string, seriesBy: string): SeriesState[] {
    const bare = sourceRange.includes("!") ? sourceRange.split("!").pop()! : sourceRange;
    const parts = bare.includes(":") ? bare.split(":") : [bare, bare];
    const start = parts[0]!;
    const end = parts[1] ?? parts[0]!;
    const col = (a: string) => {
      const m = /^([A-Z]+)/i.exec(a.replace(/\$/g, ""));
      if (!m) return 1;
      let n = 0;
      for (const ch of m[1]!.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
      return n;
    };
    const row = (a: string) => Number(/(\d+)$/.exec(a.replace(/\$/g, ""))?.[1] ?? "1");
    const cols = Math.max(1, col(end) - col(start) + 1);
    const rows = Math.max(1, row(end) - row(start) + 1);
    const mode = seriesBy === "Rows" ? "rows" : "columns";
    if (mode === "rows") {
      const count = Math.max(1, rows - 1);
      return Array.from({ length: count }, (_, i) => ({
        name: `R${i + 1}`,
        chartType: "ColumnClustered",
        smooth: false,
      }));
    }
    const count = Math.max(1, cols - 1);
    return Array.from({ length: count }, (_, i) => ({
      name: `C${i + 1}`,
      chartType: "ColumnClustered",
      smooth: false,
    }));
  }

  function makeRange(sheetName: string, address: string) {
    const bare = address.includes("!") ? address.split("!").pop()! : address;
    return {
      address: `${sheetName}!${bare.toUpperCase()}`,
      load() {},
    };
  }

  function makeWorksheet(name: string) {
    return {
      getRange(address: string) {
        return makeRange(name, address);
      },
      charts: {
        getItem(nameArg: string) {
          if (name !== chartSheetName) {
            throw new Error(`Worksheet ${name} has no charts collection access in fake`);
          }
          if (nameArg !== chart.name) throw new Error(`missing chart ${nameArg}`);
          return chartApi;
        },
      },
    };
  }

  const chartApi: {
    name: string;
    load: (props?: string) => void;
    setData?: (range: { address?: string }, seriesBy?: string) => void;
    series: {
      items: SeriesProxy[];
      load: (props?: string) => void;
    };
  } = {
    get name() {
      return chart.name;
    },
    load(_props?: string) {},
    series: {
      get items() {
        return seriesItems ?? [];
      },
      load(_props?: string) {
        const src = chart.committed;
        pendingSeriesItems = src.series.map((s) => {
          const proxy = makeSeriesProxy(s);
          (proxy as SeriesProxy & { _pendingSnap?: SeriesState })._pendingSnap = { ...s };
          return proxy;
        });
      },
    },
  };

  if (withSetData) {
    chartApi.setData = (range: { address?: string }, seriesBy?: string) => {
      const address = String(range.address ?? "A1");
      const bang = address.indexOf("!");
      const sourceSheet = bang >= 0 ? address.slice(0, bang) : chartSheetName;
      const bare = bang >= 0 ? address.slice(bang + 1) : address;
      const by = seriesBy ?? "Auto";
      chart.pending = {
        sourceRange: bare.toUpperCase(),
        seriesBy: by,
        sourceSheet,
        series: buildSeriesFromSource(bare, by),
      };
    };
  }

  const context = {
    workbook: {
      worksheets: {
        getItem(name: string) {
          const hit = sheetNames.find((s) => s.toLowerCase() === name.toLowerCase());
          if (!hit) throw new Error(`ItemNotFound: Worksheet ${name} not found`);
          return makeWorksheet(hit);
        },
      },
    },
    async sync() {
      if (chart.pending) {
        chart.committed = chart.pending;
        chart.pending = undefined;
      }
      if (pendingSeriesItems) {
        for (const proxy of pendingSeriesItems) {
          const pending = (proxy as SeriesProxy & { _pendingSnap?: SeriesState })._pendingSnap;
          if (pending) {
            seriesSnapshots.set(proxy, pending);
            delete (proxy as SeriesProxy & { _pendingSnap?: SeriesState })._pendingSnap;
          }
        }
        seriesItems = pendingSeriesItems;
        pendingSeriesItems = null;
      }
    },
  };

  (globalThis as unknown as { window: unknown }).window = globalThis;
  (globalThis as unknown as {
    Excel: { run: Function; ChartType: Record<string, string> };
  }).Excel = {
    run: async <T>(fn: (ctx: typeof context) => Promise<T>) => fn(context),
    ChartType: {
      columnClustered: "ColumnClustered",
      line: "Line",
      barClustered: "BarClustered",
      area: "Area",
      pie: "Pie",
      xyscatter: "XYScatter",
      doughnut: "Doughnut",
      bubble: "Bubble",
      radar: "Radar",
      lineMarkers: "LineMarkers",
    },
  };

  return {
    getCommitted() {
      return chart.committed;
    },
    getPending() {
      return chart.pending;
    },
    getItemsVisible() {
      return seriesItems != null;
    },
    /**
     * Broken production path: setData → load → single sync (no first sync).
     */
    async brokenUpdateSkipFirstSync(sourceRange: string) {
      const sheet = context.workbook.worksheets.getItem(chartSheetName);
      const api = sheet.charts.getItem(chart.name);
      if (!api.setData) throw new Error("setData missing");
      api.setData(sheet.getRange(sourceRange), "Columns");
      api.series.load("items/name,items/chartType,items/smooth");
      await context.sync();
      return api.series.items.map((item) => ({
        name: item.name,
        chartType: item.chartType,
        smooth: item.smooth,
      }));
    },
  };
}
