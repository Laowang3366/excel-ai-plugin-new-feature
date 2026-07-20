/**
 * Sync-gated fake for Chart.axes.
 * Writes pending until sync; load captures committed snapshot at load time.
 */

type AxisState = {
  type: string;
  axisGroup: string;
  title: string;
  titleVisible: boolean;
  minimum: number | string | null;
  maximum: number | string | null;
  majorUnit: number | string | null;
  numberFormat: string | null;
  reverse: boolean | null;
  displayUnit: string | null;
  customDisplayUnit: number | null;
  scaleType: string | null;
  logBase: number | null;
  showDisplayUnitLabel: boolean | null;
  majorGridlinesVisible: boolean | null;
  minorGridlinesVisible: boolean | null;
};

type AxisEntry = {
  committed: AxisState;
  pending: Partial<AxisState> | undefined;
};

function defaultAxis(type: string, axisGroup: string): AxisState {
  return {
    type,
    axisGroup,
    title: "",
    titleVisible: false,
    minimum: 0,
    maximum: 100,
    majorUnit: 10,
    numberFormat: "General",
    reverse: false,
    displayUnit: "None",
    customDisplayUnit: 0,
    scaleType: "Linear",
    logBase: 10,
    showDisplayUnitLabel: false,
    majorGridlinesVisible: true,
    minorGridlinesVisible: false,
  };
}

export function installChartAxesExcel(options?: { excelApi17?: boolean }) {
  const excelApi17 = options?.excelApi17 !== false;
  const axes = new Map<string, AxisEntry>();
  for (const kind of ["Category", "Value"]) {
    for (const group of ["Primary", "Secondary"]) {
      axes.set(`${kind}:${group}`, {
        committed: defaultAxis(kind, group),
        pending: undefined,
      });
    }
  }

  type AxisProxy = {
    type: string;
    axisGroup: string;
    minimum: number | string | null;
    maximum: number | string | null;
    majorUnit: number | string | null;
    numberFormat: string | null;
    reversePlotOrder: boolean | null;
    displayUnit: string | null;
    customDisplayUnit: number | null;
    scaleType: string | null;
    logBase: number | null;
    showDisplayUnitLabel: boolean | null;
    title: { text: string; visible: boolean; load: (p?: string) => void };
    majorGridlines: { visible: boolean; load: (p?: string) => void };
    minorGridlines: { visible: boolean; load: (p?: string) => void };
    setCustomDisplayUnit: (value: number) => void;
    load: (p?: string) => void;
    _flushLoad: () => void;
  };

  function makeAxisProxy(key: string): AxisProxy {
    const entry = () => {
      const found = axes.get(key);
      if (!found) throw new Error(`missing axis ${key}`);
      return found;
    };
    let pendingSnap: AxisState | undefined;
    let snapshot: AxisState | undefined;

    const proxy: AxisProxy = {
      get type() {
        if (!snapshot) throw new Error("ChartAxis.type not loaded");
        return snapshot.type;
      },
      get axisGroup() {
        if (!snapshot) throw new Error("ChartAxis.axisGroup not loaded");
        return snapshot.axisGroup;
      },
      get minimum() {
        if (!snapshot) throw new Error("ChartAxis.minimum not loaded");
        return snapshot.minimum;
      },
      set minimum(v: number | string | null) {
        entry().pending = { ...entry().pending, minimum: v };
      },
      get maximum() {
        if (!snapshot) throw new Error("ChartAxis.maximum not loaded");
        return snapshot.maximum;
      },
      set maximum(v: number | string | null) {
        entry().pending = { ...entry().pending, maximum: v };
      },
      get majorUnit() {
        if (!snapshot) throw new Error("ChartAxis.majorUnit not loaded");
        return snapshot.majorUnit;
      },
      set majorUnit(v: number | string | null) {
        entry().pending = { ...entry().pending, majorUnit: v };
      },
      get numberFormat() {
        if (!snapshot) throw new Error("ChartAxis.numberFormat not loaded");
        return snapshot.numberFormat;
      },
      set numberFormat(v: string | null) {
        entry().pending = { ...entry().pending, numberFormat: v };
      },
      get reversePlotOrder() {
        if (!snapshot) throw new Error("ChartAxis.reversePlotOrder not loaded");
        return snapshot.reverse;
      },
      set reversePlotOrder(v: boolean | null) {
        entry().pending = { ...entry().pending, reverse: v };
      },
      get displayUnit() {
        if (!snapshot) throw new Error("ChartAxis.displayUnit not loaded");
        return snapshot.displayUnit;
      },
      set displayUnit(v: string | null) {
        entry().pending = { ...entry().pending, displayUnit: v };
      },
      get customDisplayUnit() {
        if (!snapshot) throw new Error("ChartAxis.customDisplayUnit not loaded");
        return snapshot.customDisplayUnit;
      },
      get scaleType() {
        if (!snapshot) throw new Error("ChartAxis.scaleType not loaded");
        return snapshot.scaleType;
      },
      set scaleType(v: string | null) {
        entry().pending = { ...entry().pending, scaleType: v };
      },
      get logBase() {
        if (!snapshot) throw new Error("ChartAxis.logBase not loaded");
        return snapshot.logBase;
      },
      set logBase(v: number | null) {
        entry().pending = { ...entry().pending, logBase: v };
      },
      get showDisplayUnitLabel() {
        if (!snapshot) throw new Error("ChartAxis.showDisplayUnitLabel not loaded");
        return snapshot.showDisplayUnitLabel;
      },
      set showDisplayUnitLabel(v: boolean | null) {
        entry().pending = { ...entry().pending, showDisplayUnitLabel: v };
      },
      setCustomDisplayUnit(value: number) {
        entry().pending = {
          ...entry().pending,
          displayUnit: "Custom",
          customDisplayUnit: value,
        };
      },
      title: {
        get text() {
          if (!snapshot) throw new Error("ChartAxis.title.text not loaded");
          return snapshot.title;
        },
        set text(v: string) {
          entry().pending = { ...entry().pending, title: v };
        },
        get visible() {
          if (!snapshot) throw new Error("ChartAxis.title.visible not loaded");
          return snapshot.titleVisible;
        },
        set visible(v: boolean) {
          entry().pending = { ...entry().pending, titleVisible: v };
        },
        load() {},
      },
      majorGridlines: {
        get visible() {
          if (!snapshot) throw new Error("ChartAxis.majorGridlines.visible not loaded");
          return snapshot.majorGridlinesVisible as boolean;
        },
        set visible(v: boolean) {
          entry().pending = { ...entry().pending, majorGridlinesVisible: v };
        },
        load() {},
      },
      minorGridlines: {
        get visible() {
          if (!snapshot) throw new Error("ChartAxis.minorGridlines.visible not loaded");
          return snapshot.minorGridlinesVisible as boolean;
        },
        set visible(v: boolean) {
          entry().pending = { ...entry().pending, minorGridlinesVisible: v };
        },
        load() {},
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
    };
    return proxy;
  }

  const proxies = new Map<string, AxisProxy>();
  let chartNameValue: unknown = "C1";

  function proxyFor(type: string, group: string): AxisProxy {
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
        .getItem("Sheet1")
        .charts.getItem("C1")
        .axes.getItem("Value", "Primary");
      axis.minimum = 5;
      axis.maximum = 50;
      axis.load(
        "minimum,maximum,majorUnit,numberFormat,reversePlotOrder,displayUnit,customDisplayUnit,scaleType,logBase,showDisplayUnitLabel",
      );
      axis.title.load("text,visible");
      axis.majorGridlines.load("visible");
      axis.minorGridlines.load("visible");
      await context.sync();
      return { minimum: axis.minimum, maximum: axis.maximum };
    },
  };
}
