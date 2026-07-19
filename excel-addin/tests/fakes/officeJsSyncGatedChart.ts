/**
 * Chart create/list readback state machine for sync-gated Excel fake.
 * Writes stay pending until sync; load captures committed snapshot only.
 */

export type ChartState = {
  name: string;
  chartType: string;
  title: string;
  titleVisible: boolean;
  style: number;
  legendVisible: boolean;
  left: number;
  top: number;
  width: number;
  height: number;
};

export type ChartEntry = {
  committed: ChartState;
  pending: Partial<ChartState> | undefined;
};

type ChartSnap = ChartState;

type ChartProxy = {
  name: string;
  chartType: string;
  style: number;
  left: number;
  top: number;
  width: number;
  height: number;
  title: {
    text: string;
    visible: boolean;
    load: () => void;
  };
  legend: {
    visible: boolean;
    load: () => void;
  };
  load: () => void;
  delete: () => void;
  _flushLoad: () => void;
  _entry: ChartEntry;
  _sheetName: string;
};

export function createSyncGatedChartHost(deps: {
  getSheet: (name: string) => { charts: Map<string, ChartState> };
  sheets: Map<string, { charts: Map<string, ChartState> }>;
}) {
  const pendingCreates = new Map<string, Map<string, ChartEntry>>();
  const chartProxies = new Map<string, ChartProxy>();
  const chartSnapshots = new Map<string, unknown[]>();
  const pendingChartSnapshots = new Map<string, unknown[]>();

  function chartKey(sheetName: string, chartName: string) {
    return `${sheetName}\0${chartName}`;
  }

  function makeChartProxy(sheetName: string, entry: ChartEntry): ChartProxy {
    let pendingSnap: ChartSnap | undefined;
    let snapshot: ChartSnap | undefined;
    let pendingTitleSnap: { text: string; visible: boolean } | undefined;
    let titleSnap: { text: string; visible: boolean } | undefined;
    let pendingLegendSnap: boolean | undefined;
    let legendSnap: boolean | undefined;

    return {
      get name() {
        if (!snapshot) throw new Error("Chart.name not loaded");
        return snapshot.name;
      },
      set name(next: string) {
        entry.pending = { ...entry.pending, name: next };
      },
      get chartType() {
        if (!snapshot) throw new Error("Chart.chartType not loaded");
        return snapshot.chartType;
      },
      set chartType(v: string) {
        entry.pending = { ...entry.pending, chartType: v };
      },
      get style() {
        if (!snapshot) throw new Error("Chart.style not loaded");
        return snapshot.style;
      },
      set style(v: number) {
        entry.pending = { ...entry.pending, style: v };
      },
      get left() {
        if (!snapshot) throw new Error("Chart.left not loaded");
        return snapshot.left;
      },
      set left(v: number) {
        entry.pending = { ...entry.pending, left: v };
      },
      get top() {
        if (!snapshot) throw new Error("Chart.top not loaded");
        return snapshot.top;
      },
      set top(v: number) {
        entry.pending = { ...entry.pending, top: v };
      },
      get width() {
        if (!snapshot) throw new Error("Chart.width not loaded");
        return snapshot.width;
      },
      set width(v: number) {
        entry.pending = { ...entry.pending, width: v };
      },
      get height() {
        if (!snapshot) throw new Error("Chart.height not loaded");
        return snapshot.height;
      },
      set height(v: number) {
        entry.pending = { ...entry.pending, height: v };
      },
      title: {
        get text() {
          if (!titleSnap) throw new Error("Chart.title.text not loaded");
          return titleSnap.text;
        },
        set text(v: string) {
          entry.pending = { ...entry.pending, title: v };
        },
        get visible() {
          if (!titleSnap) throw new Error("Chart.title.visible not loaded");
          return titleSnap.visible;
        },
        set visible(v: boolean) {
          entry.pending = { ...entry.pending, titleVisible: v };
        },
        load() {
          pendingTitleSnap = {
            text: entry.committed.title,
            visible: entry.committed.titleVisible,
          };
          titleSnap = undefined;
        },
      },
      legend: {
        get visible() {
          if (legendSnap === undefined) throw new Error("Chart.legend.visible not loaded");
          return legendSnap;
        },
        set visible(v: boolean) {
          entry.pending = { ...entry.pending, legendVisible: v };
        },
        load() {
          pendingLegendSnap = entry.committed.legendVisible;
          legendSnap = undefined;
        },
      },
      load() {
        pendingSnap = { ...entry.committed };
        snapshot = undefined;
      },
      delete() {
        const sheet = deps.getSheet(sheetName);
        sheet.charts.delete(entry.committed.name);
        pendingCreates.get(sheetName)?.delete(entry.committed.name);
        chartProxies.delete(chartKey(sheetName, entry.committed.name));
      },
      _flushLoad() {
        if (pendingSnap) {
          snapshot = pendingSnap;
          pendingSnap = undefined;
        }
        if (pendingTitleSnap) {
          titleSnap = pendingTitleSnap;
          pendingTitleSnap = undefined;
        }
        if (pendingLegendSnap !== undefined) {
          legendSnap = pendingLegendSnap;
          pendingLegendSnap = undefined;
        }
      },
      _entry: entry,
      _sheetName: sheetName,
    };
  }

  function chartProxy(sheetName: string, chartName: string): ChartProxy {
    const key = chartKey(sheetName, chartName);
    let proxy = chartProxies.get(key);
    if (proxy) return proxy;
    const sheet = deps.getSheet(sheetName);
    const chart = sheet.charts.get(chartName);
    if (!chart) throw new Error("missing chart");
    const entry: ChartEntry = { committed: chart, pending: undefined };
    proxy = makeChartProxy(sheetName, entry);
    chartProxies.set(key, proxy);
    return proxy;
  }

  function commitPendingWrites() {
    for (const [sheetName, pendingMap] of pendingCreates) {
      const sheet = deps.getSheet(sheetName);
      for (const [name, entry] of pendingMap) {
        if (entry.pending) {
          Object.assign(entry.committed, entry.pending);
          entry.pending = undefined;
        }
        const finalName = entry.committed.name;
        sheet.charts.set(finalName, entry.committed);
        if (finalName !== name) pendingMap.delete(name);
      }
      pendingCreates.delete(sheetName);
    }
    for (const proxy of chartProxies.values()) {
      const entry = proxy._entry;
      const sheetName = proxy._sheetName;
      if (entry.pending) {
        const oldName = entry.committed.name;
        Object.assign(entry.committed, entry.pending);
        entry.pending = undefined;
        const sheet = deps.sheets.get(sheetName);
        if (sheet && entry.committed.name !== oldName) {
          sheet.charts.delete(oldName);
          sheet.charts.set(entry.committed.name, entry.committed);
          chartProxies.delete(chartKey(sheetName, oldName));
          chartProxies.set(chartKey(sheetName, entry.committed.name), proxy);
        }
      }
    }
  }

  function flushLoads() {
    for (const proxy of chartProxies.values()) {
      proxy._flushLoad();
    }
  }

  function makeChartsApi(sheetName: string) {
    const sheet = deps.getSheet(sheetName);
    return {
      get items() {
        return (chartSnapshots.get(sheetName) as unknown[]) ?? [];
      },
      load() {
        pendingChartSnapshots.set(
          sheetName,
          [...sheet.charts.keys()].map((chartName) => {
            const proxy = chartProxy(sheetName, chartName);
            proxy.load();
            proxy.title.load();
            proxy.legend.load();
            return proxy;
          }),
        );
      },
      add(type: string, _source: unknown) {
        const pendingCount = pendingCreates.get(sheetName)?.size ?? 0;
        const name = `Chart${sheet.charts.size + pendingCount + 1}`;
        const committed: ChartState = {
          name,
          chartType: "ColumnClustered",
          title: "",
          titleVisible: false,
          style: 2,
          legendVisible: true,
          left: 0,
          top: 0,
          width: 360,
          height: 240,
        };
        const entry: ChartEntry = {
          committed,
          pending: { chartType: type },
        };
        let map = pendingCreates.get(sheetName);
        if (!map) {
          map = new Map();
          pendingCreates.set(sheetName, map);
        }
        map.set(name, entry);
        const proxy = makeChartProxy(sheetName, entry);
        chartProxies.set(chartKey(sheetName, name), proxy);
        return proxy;
      },
      getItem(chartName: string) {
        if (!sheet.charts.has(chartName)) throw new Error("missing chart");
        return chartProxy(sheetName, chartName);
      },
    };
  }

  function applyChartSnapshots() {
    for (const [sheetName, pending] of pendingChartSnapshots) {
      chartSnapshots.set(sheetName, pending);
    }
    pendingChartSnapshots.clear();
  }

  return {
    commitPendingWrites,
    flushLoads,
    applyChartSnapshots,
    makeChartsApi,
    chartProxy,
  };
}
