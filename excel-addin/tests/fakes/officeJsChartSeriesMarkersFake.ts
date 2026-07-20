/**
 * Sync-gated fake for ChartSeries marker props (ExcelApi 1.7).
 * Writes queue until context.sync; property reads require load+sync.
 */

type MarkerState = {
  markerStyle: string;
  markerSize: number;
  markerBackgroundColor: string;
  markerForegroundColor: string;
};

type SeriesEntry = {
  committed: MarkerState;
  pending?: Partial<MarkerState>;
  loadQueued?: boolean;
  snapshot?: MarkerState;
};

export function installChartSeriesMarkersExcel(options?: {
  excelApi17?: boolean;
  seriesCount?: number;
}): {
  getCommitted(index0: number): MarkerState | undefined;
  poison(index0: number, patch: Partial<MarkerState>): void;
  /** After next write commit, force host style to this token before load snapshot. */
  coerceStyleAfterWrite(hostStyle: string | null): void;
  setChartName(v: unknown): void;
} {
  const excelApi17 = options?.excelApi17 !== false;
  const seriesCount = options?.seriesCount ?? 2;
  let chartNameValue: unknown = "HostChart";
  let coerceStyle: string | null = null;

  const entries: SeriesEntry[] = Array.from({ length: seriesCount }, () => ({
    committed: {
      markerStyle: "Automatic",
      markerSize: 7,
      markerBackgroundColor: "#FFFFFF",
      markerForegroundColor: "#000000",
    },
  }));

  function makeProxy(entry: SeriesEntry) {
    return {
      get markerStyle() {
        if (!entry.snapshot) throw new Error("ChartSeries.markerStyle not loaded");
        return entry.snapshot.markerStyle;
      },
      set markerStyle(v: string) {
        entry.pending = { ...entry.pending, markerStyle: v };
      },
      get markerSize() {
        if (!entry.snapshot) throw new Error("ChartSeries.markerSize not loaded");
        return entry.snapshot.markerSize;
      },
      set markerSize(v: number) {
        entry.pending = { ...entry.pending, markerSize: v };
      },
      get markerBackgroundColor() {
        if (!entry.snapshot) throw new Error("ChartSeries.markerBackgroundColor not loaded");
        return entry.snapshot.markerBackgroundColor;
      },
      set markerBackgroundColor(v: string) {
        entry.pending = { ...entry.pending, markerBackgroundColor: v };
      },
      get markerForegroundColor() {
        if (!entry.snapshot) throw new Error("ChartSeries.markerForegroundColor not loaded");
        return entry.snapshot.markerForegroundColor;
      },
      set markerForegroundColor(v: string) {
        entry.pending = { ...entry.pending, markerForegroundColor: v };
      },
      load(_props: string) {
        entry.loadQueued = true;
      },
    };
  }

  const proxies = entries.map((e) => makeProxy(e));

  const context = {
    workbook: {
      worksheets: {
        getItem(_name: string) {
          return {
            charts: {
              getItem(_chartName: string) {
                return {
                  get name() {
                    return chartNameValue;
                  },
                  load() {},
                  series: {
                    getItemAt(index: number) {
                      if (index < 0 || index >= proxies.length) {
                        throw new Error(`series index out of range: ${index}`);
                      }
                      return proxies[index]!;
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
      for (const entry of entries) {
        if (entry.pending) {
          entry.committed = { ...entry.committed, ...entry.pending };
          entry.pending = undefined;
          if (coerceStyle != null) {
            entry.committed.markerStyle = coerceStyle;
            coerceStyle = null;
          }
        }
      }
      for (const entry of entries) {
        if (entry.loadQueued) {
          entry.snapshot = { ...entry.committed };
          entry.loadQueued = false;
        }
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
          if (minVersion === "1.7") return excelApi17;
          return true;
        },
      },
    },
  };

  return {
    getCommitted(index0: number) {
      return entries[index0] ? { ...entries[index0]!.committed } : undefined;
    },
    poison(index0: number, patch: Partial<MarkerState>) {
      const e = entries[index0];
      if (!e) throw new Error("no series");
      e.committed = { ...e.committed, ...patch };
    },
    coerceStyleAfterWrite(hostStyle: string | null) {
      coerceStyle = hostStyle;
    },
    setChartName(v: unknown) {
      chartNameValue = v;
    },
  };
}
