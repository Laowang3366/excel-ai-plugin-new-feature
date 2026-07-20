import { withExcel } from "./officeJsRuntime";
import type {
  ChartTrendlineAddInput,
  ChartTrendlineDeleteResult,
  ChartTrendlineInfo,
  ChartTrendlineListResult,
  ChartTrendlineType,
  ChartTrendlineUpdateInput,
} from "./chartSeriesTrendlineTypes";
import type { HostResult } from "./types";
import { unsupported } from "./types";

const REQUIREMENT_SET = "ExcelApi";
const REQ_17 = "1.7";
const REQ_18 = "1.8";
const EVIDENCE_17 =
  "ChartSeries.trendlines / ChartTrendlineCollection.add/getItem/delete require ExcelApi 1.7";
const EVIDENCE_18 =
  "ChartTrendline forwardPeriod/backwardPeriod/showEquation/showRSquared require ExcelApi 1.8";

const TYPE_TO_HOST: Record<ChartTrendlineType, string> = {
  linear: "Linear",
  exponential: "Exponential",
  logarithmic: "Logarithmic",
  movingAverage: "MovingAverage",
  polynomial: "Polynomial",
  power: "Power",
};

type TrendlineSurface = {
  type: string;
  name: string | null;
  intercept: number | string | null;
  polynomialOrder: number | null;
  movingAveragePeriod: number | null;
  forwardPeriod?: number | null;
  backwardPeriod?: number | null;
  showEquation?: boolean | null;
  showRSquared?: boolean | null;
  delete(): void;
  load(props: string): void;
};

type TrendlineCollection = {
  items: TrendlineSurface[];
  add(type?: string): TrendlineSurface;
  getItem(index: number): TrendlineSurface;
  getCount(): { value: number };
  load(props: string): void;
};

type SeriesSurface = {
  trendlines: TrendlineCollection;
};

function isSetSupported(version: string): boolean {
  const office = (globalThis as unknown as {
    Office?: {
      context?: {
        requirements?: { isSetSupported?: (name: string, minVersion?: string) => boolean };
      };
    };
  }).Office;
  const fn = office?.context?.requirements?.isSetSupported;
  if (typeof fn !== "function") return false;
  try {
    return fn.call(office!.context!.requirements, REQUIREMENT_SET, version);
  } catch {
    return false;
  }
}

function requireLoadedString(value: unknown, field: string): string {
  if (typeof value !== "string") throw new Error(`${field} is not a loaded string`);
  return value;
}

function readString(value: unknown, field: string): string | null {
  if (value === null) return null;
  if (value === undefined) throw new Error(`${field} is not loaded`);
  if (typeof value === "string") return value;
  throw new Error(`${field} has invalid loaded type`);
}

function readNumber(value: unknown, field: string): number | null {
  if (value === null) return null;
  if (value === undefined) throw new Error(`${field} is not loaded`);
  if (typeof value === "number" && Number.isFinite(value)) return value;
  throw new Error(`${field} has invalid loaded type`);
}

function readBoolean(value: unknown, field: string): boolean | null {
  if (value === null) return null;
  if (value === undefined) throw new Error(`${field} is not loaded`);
  if (typeof value === "boolean") return value;
  throw new Error(`${field} has invalid loaded type`);
}

function mapType(raw: unknown): ChartTrendlineType | string {
  if (typeof raw !== "string") throw new Error("ChartTrendline.type is not a loaded string");
  const key = raw.toLowerCase().replace(/[^a-z]/g, "");
  const table: Record<string, ChartTrendlineType> = {
    linear: "linear",
    exponential: "exponential",
    logarithmic: "logarithmic",
    movingaverage: "movingAverage",
    polynomial: "polynomial",
    power: "power",
  };
  return table[key] ?? raw;
}

function baseLoadProps(): string {
  return "type,name,intercept,polynomialOrder,movingAveragePeriod";
}

function fullLoadProps(include18: boolean): string {
  return include18
    ? `${baseLoadProps()},forwardPeriod,backwardPeriod,showEquation,showRSquared`
    : baseLoadProps();
}

function toInfo(
  item: TrendlineSurface,
  sheetName: string,
  chartName: string,
  seriesIndex: number,
  trendlineIndex: number,
  include18: boolean,
): ChartTrendlineInfo {
  return {
    sheetName,
    chartName,
    seriesIndex,
    trendlineIndex,
    type: mapType(item.type),
    name: readString(item.name, "ChartTrendline.name"),
    // Office.js: returned intercept is always a number after load.
    intercept: readNumber(item.intercept, "ChartTrendline.intercept"),
    polynomialOrder: readNumber(item.polynomialOrder, "ChartTrendline.polynomialOrder"),
    movingAveragePeriod: readNumber(
      item.movingAveragePeriod,
      "ChartTrendline.movingAveragePeriod",
    ),
    forwardPeriod: include18
      ? readNumber(item.forwardPeriod, "ChartTrendline.forwardPeriod")
      : null,
    backwardPeriod: include18
      ? readNumber(item.backwardPeriod, "ChartTrendline.backwardPeriod")
      : null,
    showEquation: include18
      ? readBoolean(item.showEquation, "ChartTrendline.showEquation")
      : null,
    showRSquared: include18
      ? readBoolean(item.showRSquared, "ChartTrendline.showRSquared")
      : null,
  };
}

function needs18(fields: {
  forwardPeriod?: number;
  backwardPeriod?: number;
  showEquation?: boolean;
  showRSquared?: boolean;
}): boolean {
  return (
    fields.forwardPeriod !== undefined ||
    fields.backwardPeriod !== undefined ||
    fields.showEquation !== undefined ||
    fields.showRSquared !== undefined
  );
}

function applyFields(
  tl: TrendlineSurface,
  fields: {
    type?: ChartTrendlineType;
    name?: string;
    intercept?: number | "";
    polynomialOrder?: number;
    movingAveragePeriod?: number;
    forwardPeriod?: number;
    backwardPeriod?: number;
    showEquation?: boolean;
    showRSquared?: boolean;
  },
  include18: boolean,
): void {
  if (fields.type !== undefined) tl.type = TYPE_TO_HOST[fields.type];
  if (fields.name !== undefined) tl.name = fields.name;
  if (fields.intercept !== undefined) tl.intercept = fields.intercept;
  if (fields.polynomialOrder !== undefined) tl.polynomialOrder = fields.polynomialOrder;
  if (fields.movingAveragePeriod !== undefined) {
    tl.movingAveragePeriod = fields.movingAveragePeriod;
  }
  if (include18) {
    if (fields.forwardPeriod !== undefined) tl.forwardPeriod = fields.forwardPeriod;
    if (fields.backwardPeriod !== undefined) tl.backwardPeriod = fields.backwardPeriod;
    if (fields.showEquation !== undefined) tl.showEquation = fields.showEquation;
    if (fields.showRSquared !== undefined) tl.showRSquared = fields.showRSquared;
  }
}

function getSeries(
  context: { workbook: { worksheets: { getItem(name: string): unknown } } },
  sheetName: string,
  chartName: string,
  seriesIndex: number,
): { chart: { name: string; load(p: string): void }; series: SeriesSurface } {
  const sheet = context.workbook.worksheets.getItem(sheetName) as {
    charts: { getItem(name: string): { name: string; series: { getItemAt(i: number): SeriesSurface }; load(p: string): void } };
  };
  const chart = sheet.charts.getItem(chartName);
  const series = chart.series.getItemAt(seriesIndex - 1);
  return { chart, series };
}

function gate17(capability: string): HostResult<never> | null {
  if (!isSetSupported(REQ_17)) {
    return unsupported(
      capability,
      "office-js",
      "ExcelApi 1.7 is not supported in this host (Office.context.requirements.isSetSupported)",
      EVIDENCE_17,
    ) as HostResult<never>;
  }
  return null;
}

function gate18IfNeeded(
  capability: string,
  fields: {
    forwardPeriod?: number;
    backwardPeriod?: number;
    showEquation?: boolean;
    showRSquared?: boolean;
  },
): HostResult<never> | null {
  if (!needs18(fields)) return null;
  if (!isSetSupported(REQ_18)) {
    return unsupported(
      capability,
      "office-js",
      "ExcelApi 1.8 is not supported for trendline period/equation fields",
      EVIDENCE_18,
    ) as HostResult<never>;
  }
  return null;
}

/** List series trendlines; load items once then sync. */
export async function officeJsListChartSeriesTrendlines(
  sheetName: string,
  chartName: string,
  seriesIndex: number,
): Promise<HostResult<ChartTrendlineListResult>> {
  const g = gate17("chart.series.trendlines.list");
  if (g) return g;
  const include18 = isSetSupported(REQ_18);
  return withExcel("chart.series.trendlines.list", async (context) => {
    const { chart, series } = getSeries(context, sheetName, chartName, seriesIndex);
    chart.load("name");
    const props = fullLoadProps(include18)
      .split(",")
      .map((p) => `items/${p}`)
      .join(",");
    series.trendlines.load(props);
    await context.sync();
    const chartNameLoaded = requireLoadedString(chart.name, "Chart.name");
    const trendlines = series.trendlines.items.map((item, i) =>
      toInfo(item, sheetName, chartNameLoaded, seriesIndex, i + 1, include18),
    );
    return { sheetName, chartName: chartNameLoaded, seriesIndex, trendlines };
  });
}

/** Add trendline; write optional fields → sync → load new item → sync. */
export async function officeJsAddChartSeriesTrendline(
  input: ChartTrendlineAddInput,
): Promise<HostResult<ChartTrendlineInfo>> {
  const g = gate17("chart.series.trendlines.add");
  if (g) return g;
  const g18 = gate18IfNeeded("chart.series.trendlines.add", input);
  if (g18) return g18;
  const include18 = isSetSupported(REQ_18);
  return withExcel("chart.series.trendlines.add", async (context) => {
    const { chart, series } = getSeries(
      context,
      input.sheetName,
      input.chartName,
      input.seriesIndex,
    );
    // Official: ChartTrendlineCollection.add returns the new ChartTrendline proxy.
    const tl = series.trendlines.add(TYPE_TO_HOST[input.type]);
    applyFields(tl, input, include18);
    await context.sync();
    chart.load("name");
    // Load the returned object (not only collection last-item assumption).
    tl.load(fullLoadProps(include18));
    const countResult = series.trendlines.getCount();
    await context.sync();
    const chartNameLoaded = requireLoadedString(chart.name, "Chart.name");
    const count = countResult.value;
    if (typeof count !== "number" || !Number.isInteger(count) || count < 1) {
      throw new Error("ChartTrendlineCollection.getCount did not return a positive integer");
    }
    // Public trendlineIndex is 1-based; host getItem is 0-based insertion order.
    return toInfo(tl, input.sheetName, chartNameLoaded, input.seriesIndex, count, include18);
  });
}

/** Update trendline by 1-based index; write → sync → load → sync. */
export async function officeJsUpdateChartSeriesTrendline(
  input: ChartTrendlineUpdateInput,
): Promise<HostResult<ChartTrendlineInfo>> {
  const g = gate17("chart.series.trendlines.update");
  if (g) return g;
  const g18 = gate18IfNeeded("chart.series.trendlines.update", input);
  if (g18) return g18;
  const include18 = isSetSupported(REQ_18);
  return withExcel("chart.series.trendlines.update", async (context) => {
    const { chart, series } = getSeries(
      context,
      input.sheetName,
      input.chartName,
      input.seriesIndex,
    );
    const tl = series.trendlines.getItem(input.trendlineIndex - 1);
    applyFields(tl, input, include18);
    await context.sync();
    chart.load("name");
    tl.load(fullLoadProps(include18));
    await context.sync();
    return toInfo(
      tl,
      input.sheetName,
      requireLoadedString(chart.name, "Chart.name"),
      input.seriesIndex,
      input.trendlineIndex,
      include18,
    );
  });
}

/** Delete trendline; delete → sync → load remaining → sync. */
export async function officeJsDeleteChartSeriesTrendline(
  sheetName: string,
  chartName: string,
  seriesIndex: number,
  trendlineIndex: number,
): Promise<HostResult<ChartTrendlineDeleteResult>> {
  const g = gate17("chart.series.trendlines.delete");
  if (g) return g;
  const include18 = isSetSupported(REQ_18);
  return withExcel("chart.series.trendlines.delete", async (context) => {
    const { chart, series } = getSeries(context, sheetName, chartName, seriesIndex);
    series.trendlines.getItem(trendlineIndex - 1).delete();
    await context.sync();
    chart.load("name");
    const props = fullLoadProps(include18)
      .split(",")
      .map((p) => `items/${p}`)
      .join(",");
    series.trendlines.load(props);
    await context.sync();
    const chartNameLoaded = requireLoadedString(chart.name, "Chart.name");
    const remainingTrendlines = series.trendlines.items.map((item, i) =>
      toInfo(item, sheetName, chartNameLoaded, seriesIndex, i + 1, include18),
    );
    return {
      sheetName,
      chartName: chartNameLoaded,
      seriesIndex,
      deletedTrendlineIndex: trendlineIndex,
      remainingTrendlines,
    };
  });
}
