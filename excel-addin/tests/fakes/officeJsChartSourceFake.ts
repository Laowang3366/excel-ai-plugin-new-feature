/**
 * Sync-gated fake for Chart.setData + series readback.
 * setData writes stay pending until context.sync(); series.items only after load+sync.
 */

type SeriesState = { name: string; chartType: string; smooth: boolean };
type SourceState = { sourceRange: string; seriesBy: string; series: SeriesState[] };

type ChartState = {
  name: string;
  committed: SourceState;
  pending: SourceState | undefined;
};

type SeriesProxy = {
  name: string;
  chartType: string;
  smooth: boolean;
};

export function installChartSourceExcel() {
  const chart: ChartState = {
    name: "C1",
    committed: {
      sourceRange: "A1:B2",
      seriesBy: "Auto",
      series: [
        { name: "Series1", chartType: "ColumnClustered", smooth: false },
        { name: "Series2", chartType: "ColumnClustered", smooth: false },
      ],
    },
    pending: undefined,
  };

  let seriesItems: SeriesProxy[] | null = null;
  let pendingSeriesItems: SeriesProxy[] | null = null;
  const seriesSnapshots = new WeakMap<SeriesProxy, SeriesState>();

  function makeSeriesProxy(_state: SeriesState): SeriesProxy {
    const proxy: SeriesProxy = {
      get name() {
        const snap = seriesSnapshots.get(proxy);
        if (!snap) throw new Error("ChartSeries.name not loaded");
        return snap.name;
      },
      get chartType() {
        const snap = seriesSnapshots.get(proxy);
        if (!snap) throw new Error("ChartSeries.chartType not loaded");
        return snap.chartType;
      },
      get smooth() {
        const snap = seriesSnapshots.get(proxy);
        if (!snap) throw new Error("ChartSeries.smooth not loaded");
        return snap.smooth;
      },
    };
    return proxy;
  }

  function buildSeriesFromSource(sourceRange: string, seriesBy: string): SeriesState[] {
    // Minimal deterministic model: columns → 1 series per data column after header row;
    // rows → 1 series per data row after header col; auto → columns.
    const bare = sourceRange.includes("!") ? sourceRange.split("!")[1]! : sourceRange;
    const parts = bare.split(":")[0] && bare.includes(":") ? bare.split(":") : [bare, bare];
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

  const chartApi = {
    get name() {
      return chart.name;
    },
    load(_props?: string) {},
    setData(range: { address?: string }, seriesBy?: string) {
      const address = String(range.address ?? "A1");
      const bare = address.includes("!") ? address.split("!")[1]! : address;
      const by = seriesBy ?? "Auto";
      chart.pending = {
        sourceRange: bare.toUpperCase(),
        seriesBy: by,
        series: buildSeriesFromSource(bare, by),
      };
    },
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

  const context = {
    workbook: {
      worksheets: {
        getItem(name: string) {
          if (name !== "Sheet1") throw new Error(`missing sheet ${name}`);
          return {
            getRange(address: string) {
              const bare = address.includes("!") ? address.split("!")[1]! : address;
              return {
                address: `Sheet1!${bare.toUpperCase()}`,
                load() {},
              };
            },
            charts: {
              getItem(chartName: string) {
                if (chartName !== chart.name) throw new Error(`missing chart ${chartName}`);
                return chartApi;
              },
            },
          };
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
     * series load snapshots committed-at-load (old); after one sync series stay old
     * while committed source is new — proves first-sync causality.
     */
    async brokenUpdateSkipFirstSync(sourceRange: string) {
      const sheet = context.workbook.worksheets.getItem("Sheet1");
      const chartApi = sheet.charts.getItem("C1");
      chartApi.setData(sheet.getRange(sourceRange), "Columns");
      chartApi.series.load("items/name,items/chartType,items/smooth");
      await context.sync();
      return chartApi.series.items.map((item) => ({
        name: item.name,
        chartType: item.chartType,
        smooth: item.smooth,
      }));
    },
  };
}
