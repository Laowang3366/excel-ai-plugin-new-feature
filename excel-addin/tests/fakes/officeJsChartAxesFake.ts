/**
 * Sync-gated fake for Chart.axes (writes pending until sync).
 */

type AxisState = {
  type: string;
  axisGroup: string;
  title: string;
  titleVisible: boolean;
  minimum: number | string | null;
  maximum: number | string | null;
  majorUnit: number | string | null;
  minorUnit: number | string | null;
  numberFormat: string | null;
  reverse: boolean | null;
  displayUnit: string | null;
  customDisplayUnit: number | null;
  scaleType: string | null;
  logBase: number | null;
  showDisplayUnitLabel: boolean | null;
  majorGridlinesVisible: boolean | null;
  minorGridlinesVisible: boolean | null;
  majorTickMark: string | null;
  minorTickMark: string | null;
  tickLabelPosition: string | null;
  position: string | null;
  positionAt: number | null;
  linkNumberFormat: boolean | null;
};

type AxisEntry = { committed: AxisState; pending: Partial<AxisState> | undefined };

function defaultAxis(type: string, axisGroup: string): AxisState {
  return {
    type, axisGroup, title: "", titleVisible: false, minimum: 0, maximum: 100,
    majorUnit: 10, minorUnit: 2, numberFormat: "General", reverse: false,
    displayUnit: "None", customDisplayUnit: 0, scaleType: "Linear", logBase: 10,
    showDisplayUnitLabel: false, majorGridlinesVisible: true, minorGridlinesVisible: false,
    majorTickMark: "Outside", minorTickMark: "None", tickLabelPosition: "NextToAxis",
    position: "Automatic", positionAt: 0, linkNumberFormat: true,
  };
}

function resolveUnit(v: number | string | null | undefined, fallback: number) {
  return v === "" ? fallback : (v as number | string | null);
}

export function installChartAxesExcel(options?: {
  excelApi17?: boolean; excelApi18?: boolean; excelApi19?: boolean;
}) {
  const excelApi17 = options?.excelApi17 !== false;
  const excelApi18 = options?.excelApi18 !== false;
  const excelApi19 = options?.excelApi19 !== false;
  const axes = new Map<string, AxisEntry>();
  for (const kind of ["Category", "Value"]) {
    for (const group of ["Primary", "Secondary"]) {
      axes.set(`${kind}:${group}`, { committed: defaultAxis(kind, group), pending: undefined });
    }
  }

  type Proxy = {
    [k: string]: unknown;
    title: { text: string; visible: boolean; load: () => void };
    majorGridlines: { visible: boolean; load: () => void };
    minorGridlines: { visible: boolean; load: () => void };
    setCustomDisplayUnit: (v: number) => void;
    setPositionAt: (v: number) => void;
    load: () => void;
    _flushLoad: () => void;
  };

  function makeAxisProxy(key: string): Proxy {
    const entry = () => {
      const found = axes.get(key);
      if (!found) throw new Error(`missing axis ${key}`);
      return found;
    };
    let pendingSnap: AxisState | undefined;
    let snapshot: AxisState | undefined;
    const need = (field: string) => {
      if (!snapshot) throw new Error(`${field} not loaded`);
    };
    const queue = (patch: Partial<AxisState>) => {
      entry().pending = { ...entry().pending, ...patch };
    };
    const scalar = <K extends keyof AxisState>(field: string, key: K, alias?: string) => {
      const prop = alias ?? String(key);
      Object.defineProperty(proxy, prop, {
        enumerable: true,
        configurable: true,
        get() {
          need(field);
          return snapshot![key];
        },
        set(v: AxisState[K]) {
          queue({ [key]: v } as Partial<AxisState>);
        },
      });
    };
    const ro = <K extends keyof AxisState>(field: string, key: K) => {
      Object.defineProperty(proxy, key, {
        enumerable: true,
        configurable: true,
        get() {
          need(field);
          return snapshot![key];
        },
      });
    };

    const proxy = {
      title: {
        get text() {
          need("ChartAxis.title.text");
          return snapshot!.title;
        },
        set text(v: string) {
          queue({ title: v });
        },
        get visible() {
          need("ChartAxis.title.visible");
          return snapshot!.titleVisible;
        },
        set visible(v: boolean) {
          queue({ titleVisible: v });
        },
        load() {},
      },
      majorGridlines: {
        get visible() {
          need("ChartAxis.majorGridlines.visible");
          return snapshot!.majorGridlinesVisible as boolean;
        },
        set visible(v: boolean) {
          queue({ majorGridlinesVisible: v });
        },
        load() {},
      },
      minorGridlines: {
        get visible() {
          need("ChartAxis.minorGridlines.visible");
          return snapshot!.minorGridlinesVisible as boolean;
        },
        set visible(v: boolean) {
          queue({ minorGridlinesVisible: v });
        },
        load() {},
      },
      setCustomDisplayUnit(value: number) {
        queue({ displayUnit: "Custom", customDisplayUnit: value });
      },
      setPositionAt(value: number) {
        queue({ position: "Custom", positionAt: value });
      },
      load() {
        pendingSnap = { ...entry().committed };
        snapshot = undefined;
      },
      _flushLoad() {
        if (pendingSnap) {
          snapshot = pendingSnap;
          pendingSnap = undefined;
        }
      },
    } as Proxy;

    ro("ChartAxis.type", "type");
    ro("ChartAxis.axisGroup", "axisGroup");
    scalar("ChartAxis.minimum", "minimum");
    scalar("ChartAxis.maximum", "maximum");
    scalar("ChartAxis.majorUnit", "majorUnit");
    scalar("ChartAxis.minorUnit", "minorUnit");
    scalar("ChartAxis.numberFormat", "numberFormat");
    scalar("ChartAxis.reversePlotOrder", "reverse", "reversePlotOrder");
    scalar("ChartAxis.displayUnit", "displayUnit");
    ro("ChartAxis.customDisplayUnit", "customDisplayUnit");
    scalar("ChartAxis.scaleType", "scaleType");
    scalar("ChartAxis.logBase", "logBase");
    scalar("ChartAxis.showDisplayUnitLabel", "showDisplayUnitLabel");
    scalar("ChartAxis.majorTickMark", "majorTickMark");
    scalar("ChartAxis.minorTickMark", "minorTickMark");
    scalar("ChartAxis.tickLabelPosition", "tickLabelPosition");
    scalar("ChartAxis.position", "position");
    ro("ChartAxis.positionAt", "positionAt");
    scalar("ChartAxis.linkNumberFormat", "linkNumberFormat");
    return proxy;
  }

  const proxies = new Map<string, Proxy>();
  let chartNameValue: unknown = "C1";
  function proxyFor(type: string, group: string) {
    const key = `${type}:${group}`;
    let p = proxies.get(key);
    if (!p) {
      p = makeAxisProxy(key);
      proxies.set(key, p);
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
                  axes: {
                    getItem(type: string, group = "Primary") {
                      return proxyFor(type, group);
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
      for (const entry of axes.values()) {
        if (entry.pending) {
          const next = { ...entry.committed, ...entry.pending };
          next.majorUnit = resolveUnit(next.majorUnit, 10);
          next.minorUnit = resolveUnit(next.minorUnit, 2);
          entry.committed = next;
          entry.pending = undefined;
        }
      }
      for (const p of proxies.values()) p._flushLoad();
    },
  };

  (globalThis as unknown as { window: unknown }).window = globalThis;
  (globalThis as unknown as { Excel: { run: Function } }).Excel = {
    run: async <T>(fn: (ctx: typeof context) => Promise<T>) => fn(context),
  };
  (globalThis as unknown as {
    Office: { context: { requirements: { isSetSupported: (n: string, v?: string) => boolean } } };
  }).Office = {
    context: {
      requirements: {
        isSetSupported(name: string, minVersion?: string) {
          if (name !== "ExcelApi") return false;
          if (minVersion === "1.9") return excelApi19;
          if (minVersion === "1.8") return excelApi18;
          if (minVersion === "1.7") return excelApi17;
          return true;
        },
      },
    },
  };

  return {
    getCommitted(kind: string, group: string) {
      return axes.get(`${kind}:${group}`)?.committed;
    },
    getPending(kind: string, group: string) {
      return axes.get(`${kind}:${group}`)?.pending;
    },
    poisonCommitted(kind: string, group: string, patch: Partial<AxisState>) {
      const entry = axes.get(`${kind}:${group}`);
      if (!entry) throw new Error(`missing axis ${kind}:${group}`);
      entry.committed = { ...entry.committed, ...patch };
    },
    setChartName(value: unknown) {
      chartNameValue = value;
    },
    async brokenUpdateSkipFirstSync() {
      const axis = context.workbook.worksheets
        .getItem("Sheet1").charts.getItem("C1").axes.getItem("Value", "Primary");
      axis.minimum = 5;
      axis.maximum = 50;
      axis.load();
      axis.title.load();
      axis.majorGridlines.load();
      axis.minorGridlines.load();
      await context.sync();
      return { minimum: axis.minimum, maximum: axis.maximum };
    },
  };
}
