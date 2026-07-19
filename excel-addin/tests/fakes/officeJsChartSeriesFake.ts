/**
 * Sync-gated fake for Chart.series.
 * - series.items empty until load+sync
 * - writes stay pending until context.sync()
 * - load() captures a committed snapshot at call time; getters read that snapshot only after sync
 * - proves write→sync→load+sync: skipping the first sync yields old snapshot values
 */

type SeriesState = {
  name: string;
  chartType: string;
  smooth: boolean;
};

type SeriesEntry = {
  committed: SeriesState;
  pending: Partial<SeriesState> | undefined;
};

type ChartState = {
  name: string;
  series: SeriesEntry[];
  seriesItems: SeriesProxy[] | null;
  pendingSeriesItems: SeriesProxy[] | null;
};

type SheetState = {
  name: string;
  charts: Map<string, ChartState>;
};

type SeriesProxy = {
  name: string;
  chartType: string;
  smooth: boolean;
  load: (props?: string) => void;
};

export function installChartSeriesExcel(options?: {
  series?: Array<{ name: string; chartType?: string; smooth?: boolean }>;
}) {
  const initialSeries: SeriesState[] = (options?.series ?? [
    { name: "Series1", chartType: "ColumnClustered", smooth: false },
    { name: "Series2", chartType: "ColumnClustered", smooth: false },
  ]).map((s) => ({
    name: s.name,
    chartType: s.chartType ?? "ColumnClustered",
    smooth: s.smooth ?? false,
  }));

  const sheets = new Map<string, SheetState>();
  const chart: ChartState = {
    name: "C1",
    series: initialSeries.map((s) => ({
      committed: { ...s },
      pending: undefined,
    })),
    seriesItems: null,
    pendingSeriesItems: null,
  };
  sheets.set("Sheet1", {
    name: "Sheet1",
    charts: new Map([["C1", chart]]),
  });

  function makeSeriesProxy(entry: SeriesEntry): SeriesProxy {
    /** Snapshot of committed at load(); applied on next sync. */
    let pendingSnapshot: SeriesState | undefined;
    /** Readable values after load+sync. */
    let snapshot: SeriesState | undefined;

    return {
      get name() {
        if (!snapshot) {
          throw new Error("ChartSeries.name not loaded");
        }
        return snapshot.name;
      },
      set name(next: string) {
        entry.pending = { ...entry.pending, name: next };
      },
      get chartType() {
        if (!snapshot) {
          throw new Error("ChartSeries.chartType not loaded");
        }
        return snapshot.chartType;
      },
      set chartType(next: string) {
        entry.pending = { ...entry.pending, chartType: next };
      },
      get smooth() {
        if (!snapshot) {
          throw new Error("ChartSeries.smooth not loaded");
        }
        return snapshot.smooth;
      },
      set smooth(next: boolean) {
        entry.pending = { ...entry.pending, smooth: next };
      },
      load(_props?: string) {
        // Capture committed only (not pending) — requires prior sync for new writes.
        pendingSnapshot = { ...entry.committed };
        snapshot = undefined;
      },
      _flushLoad() {
        if (pendingSnapshot) {
          snapshot = pendingSnapshot;
          pendingSnapshot = undefined;
        }
      },
    } as SeriesProxy & { _flushLoad: () => void };
  }

  type ProxyWithFlush = SeriesProxy & { _flushLoad: () => void };

  function seriesCollection(chartState: ChartState) {
    const proxies = new Map<number, ProxyWithFlush>();

    function proxyAt(index: number): ProxyWithFlush {
      let proxy = proxies.get(index);
      if (!proxy) {
        proxy = makeSeriesProxy(chartState.series[index]!) as ProxyWithFlush;
        proxies.set(index, proxy);
      }
      return proxy;
    }

    return {
      get items() {
        return chartState.seriesItems ?? [];
      },
      load(_props?: string) {
        chartState.pendingSeriesItems = chartState.series.map((_, i) => {
          const proxy = proxyAt(i);
          proxy.load();
          return proxy;
        });
      },
      getItemAt(index: number) {
        if (index < 0 || index >= chartState.series.length) {
          throw new Error(`series index out of range: ${index}`);
        }
        return proxyAt(index);
      },
      _flushLoads() {
        for (const proxy of proxies.values()) {
          proxy._flushLoad();
        }
      },
    };
  }

  const collections = new Map<string, ReturnType<typeof seriesCollection>>();

  function getCollection(sheetName: string, chartName: string) {
    const key = `${sheetName}\0${chartName}`;
    let col = collections.get(key);
    if (!col) {
      const sheet = sheets.get(sheetName);
      if (!sheet) throw new Error(`missing sheet ${sheetName}`);
      const chartState = sheet.charts.get(chartName);
      if (!chartState) throw new Error(`missing chart ${chartName}`);
      col = seriesCollection(chartState);
      collections.set(key, col);
    }
    return col;
  }

  function makeChart(sheetName: string, chartName: string) {
    const sheet = sheets.get(sheetName);
    if (!sheet) throw new Error(`missing sheet ${sheetName}`);
    const chartState = sheet.charts.get(chartName);
    if (!chartState) throw new Error(`missing chart ${chartName}`);
    return {
      get name() {
        return chartState.name;
      },
      series: getCollection(sheetName, chartName),
    };
  }

  const context = {
    workbook: {
      worksheets: {
        getItem(name: string) {
          if (!sheets.has(name)) throw new Error(`missing sheet ${name}`);
          return {
            charts: {
              getItem(chartName: string) {
                return makeChart(name, chartName);
              },
            },
          };
        },
      },
    },
    async sync() {
      // 1) Commit pending writes first (Office.js order).
      for (const sheet of sheets.values()) {
        for (const chartState of sheet.charts.values()) {
          for (const entry of chartState.series) {
            if (entry.pending) {
              entry.committed = { ...entry.committed, ...entry.pending };
              entry.pending = undefined;
            }
          }
        }
      }
      // 2) Apply collection item loads queued before this sync.
      for (const sheet of sheets.values()) {
        for (const [chartName, chartState] of sheet.charts) {
          const col = collections.get(`${sheet.name}\0${chartName}`);
          if (chartState.pendingSeriesItems) {
            chartState.seriesItems = chartState.pendingSeriesItems;
            chartState.pendingSeriesItems = null;
          }
          col?._flushLoads();
        }
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
    getCommitted(seriesIndex0: number) {
      return chart.series[seriesIndex0]?.committed;
    },
    getPending(seriesIndex0: number) {
      return chart.series[seriesIndex0]?.pending;
    },
    getItemsVisible() {
      return chart.seriesItems != null;
    },
    /** Simulate broken production path: write → load → single sync (no first sync). */
    async brokenUpdateSkipFirstSync(newName: string) {
      const sheet = context.workbook.worksheets.getItem("Sheet1");
      const series = sheet.charts.getItem("C1").series.getItemAt(0);
      series.name = newName;
      series.load("name,chartType,smooth");
      await context.sync();
      return {
        name: series.name,
        chartType: series.chartType,
        smooth: series.smooth,
      };
    },
  };
}
