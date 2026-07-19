/**
 * Sync-gated fake for ChartSeries.hasDataLabels + ChartSeries.dataLabels.
 * Office.context.requirements.isSetSupported precheck must pass before writes.
 */

type LabelsState = {
  showValue: boolean;
  showCategoryName: boolean;
  showSeriesName: boolean;
  numberFormat: string;
};

type SeriesEntry = {
  committedEnabled: boolean;
  pendingEnabled: boolean | undefined;
  committed: LabelsState;
  pending: Partial<LabelsState> | undefined;
};

export function installChartDataLabelsExcel(options?: {
  seriesCount?: number;
  /** When false, isSetSupported('ExcelApi','1.7') is false. Default true. */
  excelApi17?: boolean;
  /** When false, isSetSupported('ExcelApi','1.8') is false. Default true. */
  excelApi18?: boolean;
  /** When false, series has no hasDataLabels property. Default true. */
  supportHasDataLabels?: boolean;
  /** When false, series has no dataLabels property. Default true. */
  supportDataLabels?: boolean;
  sheetNameValue?: unknown;
  chartNameValue?: unknown;
}) {
  const seriesCount = options?.seriesCount ?? 2;
  const excelApi17 = options?.excelApi17 !== false;
  const excelApi18 = options?.excelApi18 !== false;
  const supportHasDataLabels = options?.supportHasDataLabels !== false;
  const supportDataLabels = options?.supportDataLabels !== false;

  const series: SeriesEntry[] = Array.from({ length: seriesCount }, () => ({
    committedEnabled: false,
    pendingEnabled: undefined,
    committed: {
      showValue: false,
      showCategoryName: false,
      showSeriesName: false,
      numberFormat: "General",
    },
    pending: undefined,
  }));

  type LabelsProxy = {
    showValue: boolean;
    showCategoryName: boolean;
    showSeriesName: boolean;
    numberFormat: string;
    load: (p?: string) => void;
    _flushLoad: () => void;
  };

  type SeriesProxy = {
    hasDataLabels?: boolean;
    dataLabels?: LabelsProxy;
    load: (p?: string) => void;
    _flushLoad: () => void;
  };

  let sheetNameValue: unknown = options?.sheetNameValue ?? "Sheet1";
  let chartNameValue: unknown = options?.chartNameValue ?? "C1";
  let hasDataLabelsWriteCalls = 0;
  let dataLabelsWriteCalls = 0;
  let excelRunCalls = 0;
  let worksheetGetItemCalls = 0;

  function makeLabelsProxy(entry: SeriesEntry): LabelsProxy {
    let pendingSnap: LabelsState | undefined;
    let snapshot: LabelsState | undefined;
    return {
      get showValue() {
        if (!snapshot) throw new Error("ChartDataLabels.showValue not loaded");
        return snapshot.showValue;
      },
      set showValue(v: boolean) {
        dataLabelsWriteCalls += 1;
        entry.pending = { ...entry.pending, showValue: v };
      },
      get showCategoryName() {
        if (!snapshot) throw new Error("ChartDataLabels.showCategoryName not loaded");
        return snapshot.showCategoryName;
      },
      set showCategoryName(v: boolean) {
        dataLabelsWriteCalls += 1;
        entry.pending = { ...entry.pending, showCategoryName: v };
      },
      get showSeriesName() {
        if (!snapshot) throw new Error("ChartDataLabels.showSeriesName not loaded");
        return snapshot.showSeriesName;
      },
      set showSeriesName(v: boolean) {
        dataLabelsWriteCalls += 1;
        entry.pending = { ...entry.pending, showSeriesName: v };
      },
      get numberFormat() {
        if (!snapshot) throw new Error("ChartDataLabels.numberFormat not loaded");
        return snapshot.numberFormat;
      },
      set numberFormat(v: string) {
        dataLabelsWriteCalls += 1;
        entry.pending = { ...entry.pending, numberFormat: v };
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
  }

  function makeSeriesProxy(entry: SeriesEntry): SeriesProxy {
    let pendingEnabledSnap: boolean | undefined;
    let enabledSnapshot: boolean | undefined;
    const labelsProxy = supportDataLabels ? makeLabelsProxy(entry) : undefined;
    const proxy: SeriesProxy = {
      load(props?: string) {
        if (!props || props.includes("hasDataLabels")) {
          pendingEnabledSnap = entry.committedEnabled;
          enabledSnapshot = undefined;
        }
      },
      _flushLoad() {
        if (pendingEnabledSnap !== undefined) {
          enabledSnapshot = pendingEnabledSnap;
          pendingEnabledSnap = undefined;
        }
        labelsProxy?._flushLoad();
      },
    };
    if (supportHasDataLabels) {
      Object.defineProperty(proxy, "hasDataLabels", {
        enumerable: true,
        configurable: true,
        get() {
          if (enabledSnapshot === undefined) {
            throw new Error("ChartSeries.hasDataLabels not loaded");
          }
          return enabledSnapshot;
        },
        set(v: boolean) {
          hasDataLabelsWriteCalls += 1;
          entry.pendingEnabled = v;
        },
      });
    }
    if (supportDataLabels && labelsProxy) {
      Object.defineProperty(proxy, "dataLabels", {
        enumerable: true,
        configurable: true,
        get() {
          return labelsProxy;
        },
      });
    }
    return proxy;
  }

  const proxies = new Map<number, SeriesProxy>();

  function proxyAt(index: number): SeriesProxy {
    let p = proxies.get(index);
    if (!p) {
      p = makeSeriesProxy(series[index]!);
      proxies.set(index, p);
    }
    return p;
  }

  const context = {
    workbook: {
      worksheets: {
        getItem(name: string) {
          worksheetGetItemCalls += 1;
          if (name !== "Sheet1") throw new Error(`missing sheet ${name}`);
          return {
            get name() {
              return sheetNameValue as string;
            },
            load() {},
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
        if (entry.pendingEnabled !== undefined) {
          entry.committedEnabled = entry.pendingEnabled;
          entry.pendingEnabled = undefined;
        }
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
  (globalThis as unknown as {
    Office: {
      context: {
        requirements: { isSetSupported: (name: string, minVersion?: string) => boolean };
      };
    };
  }).Office = {
    context: {
      requirements: {
        isSetSupported(name: string, minVersion?: string) {
          if (name !== "ExcelApi") return false;
          if (minVersion === "1.7") return excelApi17;
          if (minVersion === "1.8") return excelApi18;
          return false;
        },
      },
    },
  };
  (globalThis as unknown as { Excel: { run: Function } }).Excel = {
    run: async <T>(fn: (ctx: typeof context) => Promise<T>) => {
      excelRunCalls += 1;
      return fn(context);
    },
  };

  return {
    getCommitted(index0: number) {
      const entry = series[index0];
      if (!entry) return undefined;
      return { enabled: entry.committedEnabled, ...entry.committed };
    },
    getPending(index0: number) {
      const entry = series[index0];
      if (!entry) return undefined;
      return { enabled: entry.pendingEnabled, ...entry.pending };
    },
    getWriteCallCounts() {
      return {
        hasDataLabelsWriteCalls,
        dataLabelsWriteCalls,
        excelRunCalls,
        worksheetGetItemCalls,
      };
    },
    poisonCommitted(index0: number, patch: Record<string, unknown>) {
      const entry = series[index0];
      if (!entry) throw new Error("missing series");
      if ("enabled" in patch) {
        entry.committedEnabled = patch.enabled as boolean;
      }
      const { enabled: _e, ...labelsPatch } = patch as { enabled?: unknown } & Partial<LabelsState>;
      entry.committed = { ...entry.committed, ...labelsPatch } as LabelsState;
    },
    setSheetName(value: unknown) {
      sheetNameValue = value;
    },
    setChartName(value: unknown) {
      chartNameValue = value;
    },
    /**
     * Stale path: write then load without first sync → load samples old committed.
     */
    async brokenUpdateSkipFirstSync() {
      const seriesProxy = context.workbook.worksheets
        .getItem("Sheet1")
        .charts.getItem("C1")
        .series.getItemAt(0);
      seriesProxy.hasDataLabels = true;
      seriesProxy.dataLabels!.showValue = true;
      seriesProxy.load("hasDataLabels");
      seriesProxy.dataLabels!.load("showValue,showCategoryName,showSeriesName,numberFormat");
      await context.sync();
      return {
        enabled: seriesProxy.hasDataLabels,
        showValue: seriesProxy.dataLabels!.showValue,
        showCategoryName: seriesProxy.dataLabels!.showCategoryName,
      };
    },
  };
}
