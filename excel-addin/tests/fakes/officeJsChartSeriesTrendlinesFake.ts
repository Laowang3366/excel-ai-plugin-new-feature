/**
 * Sync-gated fake for ChartSeries.trendlines (ExcelApi 1.7/1.8).
 */

type TrendlineState = {
  type: string;
  name: string | null;
  intercept: number | string | null;
  polynomialOrder: number | null;
  movingAveragePeriod: number | null;
  forwardPeriod: number | null;
  backwardPeriod: number | null;
  showEquation: boolean | null;
  showRSquared: boolean | null;
};

type Entry = { committed: TrendlineState; pending?: Partial<TrendlineState> };

function defaultTl(type = "Linear"): TrendlineState {
  return {
    type,
    name: null,
    intercept: "",
    polynomialOrder: 2,
    movingAveragePeriod: 2,
    forwardPeriod: 0,
    backwardPeriod: 0,
    showEquation: false,
    showRSquared: false,
  };
}

export function installChartSeriesTrendlinesExcel(options?: {
  excelApi17?: boolean;
  excelApi18?: boolean;
}) {
  const excelApi17 = options?.excelApi17 !== false;
  const excelApi18 = options?.excelApi18 !== false;
  const items: Entry[] = [];
  let chartNameValue: unknown = "C1";

  type TlProxy = {
    type: string;
    name: string | null;
    intercept: number | string | null;
    polynomialOrder: number | null;
    movingAveragePeriod: number | null;
    forwardPeriod: number | null;
    backwardPeriod: number | null;
    showEquation: boolean | null;
    showRSquared: boolean | null;
    delete(): void;
    load(p?: string): void;
    _flushLoad(): void;
    _entry: Entry;
  };

  function makeProxy(entry: Entry): TlProxy {
    let pendingSnap: TrendlineState | undefined;
    let snapshot: TrendlineState | undefined;
    const proxy: TlProxy = {
      _entry: entry,
      get type() {
        if (!snapshot) throw new Error("ChartTrendline.type not loaded");
        return snapshot.type;
      },
      set type(v: string) {
        entry.pending = { ...entry.pending, type: v };
      },
      get name() {
        if (!snapshot) throw new Error("ChartTrendline.name not loaded");
        return snapshot.name;
      },
      set name(v: string | null) {
        entry.pending = { ...entry.pending, name: v };
      },
      get intercept() {
        if (!snapshot) throw new Error("ChartTrendline.intercept not loaded");
        return snapshot.intercept;
      },
      set intercept(v: number | string | null) {
        entry.pending = { ...entry.pending, intercept: v };
      },
      get polynomialOrder() {
        if (!snapshot) throw new Error("ChartTrendline.polynomialOrder not loaded");
        return snapshot.polynomialOrder;
      },
      set polynomialOrder(v: number | null) {
        entry.pending = { ...entry.pending, polynomialOrder: v };
      },
      get movingAveragePeriod() {
        if (!snapshot) throw new Error("ChartTrendline.movingAveragePeriod not loaded");
        return snapshot.movingAveragePeriod;
      },
      set movingAveragePeriod(v: number | null) {
        entry.pending = { ...entry.pending, movingAveragePeriod: v };
      },
      get forwardPeriod() {
        if (!snapshot) throw new Error("ChartTrendline.forwardPeriod not loaded");
        return snapshot.forwardPeriod;
      },
      set forwardPeriod(v: number | null) {
        entry.pending = { ...entry.pending, forwardPeriod: v };
      },
      get backwardPeriod() {
        if (!snapshot) throw new Error("ChartTrendline.backwardPeriod not loaded");
        return snapshot.backwardPeriod;
      },
      set backwardPeriod(v: number | null) {
        entry.pending = { ...entry.pending, backwardPeriod: v };
      },
      get showEquation() {
        if (!snapshot) throw new Error("ChartTrendline.showEquation not loaded");
        return snapshot.showEquation;
      },
      set showEquation(v: boolean | null) {
        entry.pending = { ...entry.pending, showEquation: v };
      },
      get showRSquared() {
        if (!snapshot) throw new Error("ChartTrendline.showRSquared not loaded");
        return snapshot.showRSquared;
      },
      set showRSquared(v: boolean | null) {
        entry.pending = { ...entry.pending, showRSquared: v };
      },
      delete() {
        entry.pending = { ...entry.pending, type: "__deleted__" };
      },
      load() {
        pendingSnap = { ...entry.committed };
        snapshot = undefined;
      },
      _flushLoad() {
        if (pendingSnap) {
          snapshot = pendingSnap;
          pendingSnap = undefined;
        }
      },
    };
    return proxy;
  }

  let proxies: TlProxy[] = [];

  function rebuildProxies() {
    proxies = items.map((e) => makeProxy(e));
  }

  const collection = {
    get items() {
      return proxies;
    },
    add(type = "Linear") {
      const entry: Entry = { committed: defaultTl(type), pending: undefined };
      // add is pending until sync: store as pending-only via committed empty then pending type
      entry.pending = { ...defaultTl(type), type };
      // temporary committed baseline
      entry.committed = defaultTl(type);
      items.push(entry);
      rebuildProxies();
      return proxies[proxies.length - 1]!;
    },
    getItem(index: number) {
      if (index < 0 || index >= proxies.length) throw new Error(`trendline index ${index}`);
      return proxies[index]!;
    },
    load() {
      for (const p of proxies) p.load();
    },
  };

  const context = {
    workbook: {
      worksheets: {
        getItem(name: string) {
          if (name !== "Sheet1") throw new Error(`missing sheet ${name}`);
          return {
            charts: {
              getItem(chartName: string) {
                if (chartName !== "C1") throw new Error(`missing chart ${chartName}`);
                return {
                  get name() {
                    return chartNameValue as string;
                  },
                  load() {},
                  series: {
                    getItemAt(i: number) {
                      if (i !== 0) throw new Error(`series index ${i}`);
                      return { trendlines: collection };
                    },
                  },
                };
              },
            },
          };
        },
      },
    },
        async sync() {
      // Apply property writes first.
      for (const entry of items) {
        if (entry.pending && entry.pending.type !== "__deleted__") {
          entry.committed = { ...entry.committed, ...entry.pending };
          entry.pending = undefined;
        }
      }
      // Then structural deletes.
      let structural = false;
      for (let i = items.length - 1; i >= 0; i--) {
        if (items[i]!.pending?.type === "__deleted__") {
          items.splice(i, 1);
          structural = true;
        }
      }
      if (structural || proxies.length !== items.length) {
        rebuildProxies();
      }
      for (const p of proxies) {
        p._flushLoad();
      }
    },
  };

  (globalThis as unknown as { window: unknown }).window = globalThis;
  (globalThis as unknown as { Excel: { run: Function } }).Excel = {
    run: async <T>(fn: (ctx: typeof context) => Promise<T>) => fn(context),
  };
  (globalThis as unknown as {
    Office: {
      context: { requirements: { isSetSupported: (n: string, v?: string) => boolean } };
    };
  }).Office = {
    context: {
      requirements: {
        isSetSupported(name: string, minVersion?: string) {
          if (name !== "ExcelApi") return false;
          if (minVersion === "1.8") return excelApi18;
          if (minVersion === "1.7") return excelApi17;
          return true;
        },
      },
    },
  };

  rebuildProxies();

  return {
    getCommitted() {
      return items.map((e) => ({ ...e.committed }));
    },
    poisonLast(patch: Partial<TrendlineState>) {
      if (items.length === 0) throw new Error("no trendlines");
      const last = items[items.length - 1]!;
      last.committed = { ...last.committed, ...patch };
    },
    setChartName(v: unknown) {
      chartNameValue = v;
    },
  };
}
