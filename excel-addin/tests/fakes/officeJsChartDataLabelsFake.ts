/**
 * Sync-gated fake for ChartSeries.hasDataLabels + ChartSeries.dataLabels.
 */
type LabelsState = {
  showValue: boolean;
  showCategoryName: boolean;
  showSeriesName: boolean;
  numberFormat: string;
  showPercentage: boolean;
  showBubbleSize: boolean;
  showLegendKey: boolean;
  separator: string;
  position: string;
};

type SeriesEntry = {
  committedEnabled: boolean;
  pendingEnabled: boolean | undefined;
  committed: LabelsState;
  pending: Partial<LabelsState> | undefined;
};

const DEFAULT_LABELS: LabelsState = {
  showValue: false,
  showCategoryName: false,
  showSeriesName: false,
  numberFormat: "General",
  showPercentage: false,
  showBubbleSize: false,
  showLegendKey: false,
  separator: ", ",
  position: "Center",
};

const LABEL_KEYS = Object.keys(DEFAULT_LABELS) as Array<keyof LabelsState>;

export function installChartDataLabelsExcel(options?: {
  seriesCount?: number;
  excelApi17?: boolean;
  excelApi18?: boolean;
  missingIsSetSupported?: boolean;
  isSetSupportedThrows?: boolean;
  supportHasDataLabels?: boolean;
  supportDataLabels?: boolean;
  /** When false, extended 1.8 props are absent on proxy. Default true. */
  supportExtendedLabels?: boolean;
  sheetNameValue?: unknown;
  chartNameValue?: unknown;
}) {
  const seriesCount = options?.seriesCount ?? 2;
  const excelApi17 = options?.excelApi17 !== false;
  const excelApi18 = options?.excelApi18 !== false;
  const supportHasDataLabels = options?.supportHasDataLabels !== false;
  const supportDataLabels = options?.supportDataLabels !== false;
  const supportExtended = options?.supportExtendedLabels !== false;

  const series: SeriesEntry[] = Array.from({ length: seriesCount }, () => ({
    committedEnabled: false,
    pendingEnabled: undefined,
    committed: { ...DEFAULT_LABELS },
    pending: undefined,
  }));

  type LabelsProxy = LabelsState & { load: (p?: string) => void; _flushLoad: () => void };
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
    const proxy = {
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
    } as LabelsProxy;

    for (const key of LABEL_KEYS) {
      if (
        !supportExtended &&
        (key === "showPercentage" ||
          key === "showBubbleSize" ||
          key === "showLegendKey" ||
          key === "separator" ||
          key === "position")
      ) {
        continue;
      }
      Object.defineProperty(proxy, key, {
        enumerable: true,
        configurable: true,
        get() {
          if (!snapshot) throw new Error(`ChartDataLabels.${key} not loaded`);
          return snapshot[key];
        },
        set(v: LabelsState[typeof key]) {
          dataLabelsWriteCalls += 1;
          entry.pending = { ...entry.pending, [key]: v };
        },
      });
    }
    return proxy;
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
    if (labelsProxy) proxy.dataLabels = labelsProxy;
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
      for (const p of proxies.values()) p._flushLoad();
    },
  };

  (globalThis as unknown as { window: unknown }).window = globalThis;
  if (options?.missingIsSetSupported) {
    (globalThis as unknown as { Office: unknown }).Office = {
      context: { requirements: {} },
    };
  } else if (options?.isSetSupportedThrows) {
    (globalThis as unknown as { Office: unknown }).Office = {
      context: {
        requirements: {
          isSetSupported: () => {
            throw new Error("isSetSupported boom");
          },
        },
      },
    };
  } else {
    (globalThis as unknown as {
      Office: {
        context: {
          requirements: { isSetSupported: (n: string, v?: string) => boolean };
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
  }
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
      if ("enabled" in patch) entry.committedEnabled = patch.enabled as boolean;
      const { enabled: _e, ...labelsPatch } = patch as { enabled?: unknown } & Partial<LabelsState>;
      entry.committed = { ...entry.committed, ...labelsPatch } as LabelsState;
    },
    setSheetName(value: unknown) {
      sheetNameValue = value;
    },
    setChartName(value: unknown) {
      chartNameValue = value;
    },
    async brokenUpdateSkipFirstSync() {
      const seriesProxy = context.workbook.worksheets
        .getItem("Sheet1")
        .charts.getItem("C1")
        .series.getItemAt(0);
      seriesProxy.hasDataLabels = true;
      seriesProxy.dataLabels!.showValue = true;
      seriesProxy.load("hasDataLabels");
      seriesProxy.dataLabels!.load();
      await context.sync();
      return {
        enabled: seriesProxy.hasDataLabels,
        showValue: seriesProxy.dataLabels!.showValue,
        showCategoryName: seriesProxy.dataLabels!.showCategoryName,
      };
    },
  };
}
