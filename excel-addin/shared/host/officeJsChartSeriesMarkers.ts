import { withExcel } from "./officeJsRuntime";
import type {
  ChartMarkerStyle,
  ChartSeriesMarkersInfo,
  ChartSeriesMarkersUpdateInput,
} from "./chartSeriesMarkersTypes";
import type { HostResult } from "./types";
import { unsupported } from "./types";

const REQUIREMENT_SET = "ExcelApi";
const REQ_17 = "1.7";
const EVIDENCE =
  "ChartSeries.markerStyle/markerSize/markerBackgroundColor/markerForegroundColor require ExcelApi 1.7";

const STYLE_TO_HOST: Record<ChartMarkerStyle, string> = {
  automatic: "Automatic",
  none: "None",
  square: "Square",
  diamond: "Diamond",
  triangle: "Triangle",
  x: "X",
  star: "Star",
  dot: "Dot",
  dash: "Dash",
  circle: "Circle",
  plus: "Plus",
  picture: "Picture",
};

type SeriesSurface = {
  markerStyle: string;
  markerSize: number;
  markerBackgroundColor: string;
  markerForegroundColor: string;
  load(props: string): void;
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

function requireLoadedNumber(value: unknown, field: string): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error(`${field} is not a loaded finite number`);
  }
  return value;
}

/** Normalize host color string to #RRGGBB when possible. */
export function normalizeMarkerColor(raw: string, field: string): string {
  const t = raw.trim();
  const hex = t.startsWith("#") ? t.slice(1) : t;
  // Drop optional alpha if host returns #AARRGGBB / AARRGGBB (keep last 6).
  const six = hex.length === 8 ? hex.slice(2) : hex;
  if (!/^[0-9A-Fa-f]{6}$/.test(six)) {
    throw new Error(`${field} is not a #RRGGBB color`);
  }
  return `#${six.toUpperCase()}`;
}

function mapStyle(raw: unknown): ChartMarkerStyle | string {
  if (typeof raw !== "string") throw new Error("ChartSeries.markerStyle is not a loaded string");
  const key = raw.toLowerCase().replace(/[^a-z]/g, "");
  const table: Record<string, ChartMarkerStyle> = {
    automatic: "automatic",
    none: "none",
    square: "square",
    diamond: "diamond",
    triangle: "triangle",
    x: "x",
    star: "star",
    dot: "dot",
    dash: "dash",
    circle: "circle",
    plus: "plus",
    picture: "picture",
  };
  return table[key] ?? raw;
}

function toInfo(
  series: SeriesSurface,
  sheetName: string,
  chartName: string,
  seriesIndex: number,
): ChartSeriesMarkersInfo {
  return {
    sheetName,
    chartName,
    seriesIndex,
    markerStyle: mapStyle(series.markerStyle),
    markerSize: requireLoadedNumber(series.markerSize, "ChartSeries.markerSize"),
    markerBackgroundColor: normalizeMarkerColor(
      requireLoadedString(series.markerBackgroundColor, "ChartSeries.markerBackgroundColor"),
      "ChartSeries.markerBackgroundColor",
    ),
    markerForegroundColor: normalizeMarkerColor(
      requireLoadedString(series.markerForegroundColor, "ChartSeries.markerForegroundColor"),
      "ChartSeries.markerForegroundColor",
    ),
  };
}

/**
 * Compare requested fields to host readback.
 * Colors are compared after #RRGGBB normalization so direct adapter calls
 * with lowercase/unprefixed hex do not false-fail (executor also normalizes).
 */
function assertMatches(
  info: ChartSeriesMarkersInfo,
  expected: {
    markerStyle?: ChartMarkerStyle;
    markerSize?: number;
    markerBackgroundColor?: string;
    markerForegroundColor?: string;
  },
): void {
  if (expected.markerStyle !== undefined && info.markerStyle !== expected.markerStyle) {
    throw new Error(
      `markerStyle readback mismatch: expected ${expected.markerStyle}, got ${info.markerStyle}`,
    );
  }
  if (expected.markerSize !== undefined && info.markerSize !== expected.markerSize) {
    throw new Error(
      `markerSize readback mismatch: expected ${expected.markerSize}, got ${info.markerSize}`,
    );
  }
  if (
    expected.markerBackgroundColor !== undefined &&
    info.markerBackgroundColor !== expected.markerBackgroundColor
  ) {
    throw new Error(
      `markerBackgroundColor readback mismatch: expected ${expected.markerBackgroundColor}, got ${info.markerBackgroundColor}`,
    );
  }
  if (
    expected.markerForegroundColor !== undefined &&
    info.markerForegroundColor !== expected.markerForegroundColor
  ) {
    throw new Error(
      `markerForegroundColor readback mismatch: expected ${expected.markerForegroundColor}, got ${info.markerForegroundColor}`,
    );
  }
}

/** Update series markers; ExcelApi 1.7; write → sync → load → sync. */
export async function officeJsUpdateChartSeriesMarkers(
  input: ChartSeriesMarkersUpdateInput,
): Promise<HostResult<ChartSeriesMarkersInfo>> {
  if (!isSetSupported(REQ_17)) {
    return unsupported(
      "chart.series.markers.update",
      "office-js",
      "ExcelApi 1.7 is not supported in this host (Office.context.requirements.isSetSupported)",
      EVIDENCE,
    );
  }

  return withExcel("chart.series.markers.update", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName) as unknown as {
      charts: {
        getItem(name: string): {
          name: string;
          load(p: string): void;
          series: { getItemAt(index: number): SeriesSurface };
        };
      };
    };
    const chart = sheet.charts.getItem(input.chartName);
    const series = chart.series.getItemAt(input.seriesIndex - 1);
    // Normalize colors at host layer so adapter-direct and executor paths share contract.
    const markerBackgroundColor =
      input.markerBackgroundColor !== undefined
        ? normalizeMarkerColor(input.markerBackgroundColor, "markerBackgroundColor")
        : undefined;
    const markerForegroundColor =
      input.markerForegroundColor !== undefined
        ? normalizeMarkerColor(input.markerForegroundColor, "markerForegroundColor")
        : undefined;
    if (input.markerStyle !== undefined) series.markerStyle = STYLE_TO_HOST[input.markerStyle];
    if (input.markerSize !== undefined) series.markerSize = input.markerSize;
    if (markerBackgroundColor !== undefined) {
      series.markerBackgroundColor = markerBackgroundColor;
    }
    if (markerForegroundColor !== undefined) {
      series.markerForegroundColor = markerForegroundColor;
    }
    await context.sync();
    chart.load("name");
    series.load("markerStyle,markerSize,markerBackgroundColor,markerForegroundColor");
    await context.sync();
    const chartName = requireLoadedString(chart.name, "Chart.name");
    const info = toInfo(series, input.sheetName, chartName, input.seriesIndex);
    assertMatches(info, {
      markerStyle: input.markerStyle,
      markerSize: input.markerSize,
      markerBackgroundColor,
      markerForegroundColor,
    });
    return info;
  });
}
