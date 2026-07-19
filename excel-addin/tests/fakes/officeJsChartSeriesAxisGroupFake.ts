/**
 * Sync-gated fake for ChartSeries.axisGroup.
 */

type SeriesState = { axisGroup: string };
type SeriesEntry = { committed: SeriesState; pending: Partial<SeriesState> | undefined };

export function installChartSeriesAxisGroupExcel(seriesCount = 2) {
  const series: SeriesEntry[] = Array.from({ length: seriesCount }, () => ({
    committed: { axisGroup: "Primary" },
    pending: undefined,
  }));

  type SeriesProxy = {
    axisGroup: string;
    load: (p?: string) => void;
    _flushLoad: () => void;
  };

  const proxies = new Map<number, SeriesProxy>();
  let chartNameValue: unknown = "C1";
  /** When set, series load() captures this axisGroup instead of committed (post-write). */
  const loadOverrides = new Map<number, string>();

  function makeProxyWithIndex(entry: SeriesEntry, index: number): SeriesProxy {
    let pendingSnap: SeriesState | undefined;
    let snapshot: SeriesState | undefined;
    return {
      get axisGroup() {
        if (!snapshot) throw new Error("ChartSeries.axisGroup not loaded");
        return snapshot.axisGroup;
      },
      set axisGroup(v: string) {
        entry.pending = { ...entry.pending, axisGroup: v };
      },
      load() {
        const override = loadOverrides.get(index);
        pendingSnap =
          override !== undefined ? { axisGroup: override } : { ...entry.committed };
        snapshot = undefined;
      },
      _flushLoad() {
        if (pendingSnap) {
          snapshot = pendingSnap;
          pendingSnap = undefined;
        }
      },
    };
  }

  function proxyAt(index: number): SeriesProxy {
    let p = proxies.get(index);
    if (!p) {
      p = makeProxyWithIndex(series[index]!, index);
      proxies.set(index, p);
    }
    return p;
  }

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
                    getItemAt(index: number) {
                      if (index < 0 || index >= series.length) {
                        throw new Error(`series index out of range: ${index}`);
                      }
                      return proxyAt(index);
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
      for (const entry of series) {
        if (entry.pending) {
          entry.committed = { ...entry.committed, ...entry.pending };
          entry.pending = undefined;
        }
      }
      for (const p of proxies.values()) {
        p._flushLoad();
      }
    },
  };

  (globalThis as unknown as { window: unknown }).window = globalThis;
  (globalThis as unknown as { Excel: { run: Function } }).Excel = {
    run: async <T>(fn: (ctx: typeof context) => Promise<T>) => fn(context),
  };

  return {
    getCommitted(index0: number) {
      return series[index0]?.committed;
    },
    getPending(index0: number) {
      return series[index0]?.pending;
    },
    setChartName(value: unknown) {
      chartNameValue = value;
    },
    /** Force next load() snapshot to use this host string (after write sync). */
    setLoadOverride(index0: number, axisGroup: string) {
      loadOverrides.set(index0, axisGroup);
    },
    async brokenUpdateSkipFirstSync() {
      const s = context.workbook.worksheets
        .getItem("Sheet1")
        .charts.getItem("C1")
        .series.getItemAt(0);
      s.axisGroup = "Secondary";
      s.load("axisGroup");
      await context.sync();
      return { axisGroup: s.axisGroup };
    },
  };
}
