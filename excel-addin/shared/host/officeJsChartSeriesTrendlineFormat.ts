import { withExcel } from "./officeJsRuntime";
import type {
  ChartLineStyle,
  ChartTrendlineFormatInfo,
  ChartTrendlineFormatUpdateInput,
} from "./chartSeriesTrendlineFormatTypes";
import type { HostResult } from "./types";
import { unsupported } from "./types";

const REQUIREMENT_SET = "ExcelApi";
const REQ_17 = "1.7";
const EVIDENCE =
  "ChartTrendline.format.line (color/lineStyle/weight) requires ExcelApi 1.7 (format.line from 1.7; color scalar also 1.1)";

const STYLE_TO_HOST: Record<ChartLineStyle, string> = {
  none: "None",
  continuous: "Continuous",
  dash: "Dash",
  dashDot: "DashDot",
  dashDotDot: "DashDotDot",
  dot: "Dot",
  grey25: "Grey25",
  grey50: "Grey50",
  grey75: "Grey75",
  automatic: "Automatic",
  roundDot: "RoundDot",
};

type LineSurface = {
  color: string;
  lineStyle: string;
  weight: number;
  load(props: string): void;
};

type TrendlineSurface = {
  format: { line: LineSurface };
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

/** Normalize host color string to #RRGGBB. */
export function normalizeTrendlineLineColor(raw: string, field: string): string {
  const t = raw.trim();
  const hex = t.startsWith("#") ? t.slice(1) : t;
  const six = hex.length === 8 ? hex.slice(2) : hex;
  if (!/^[0-9A-Fa-f]{6}$/.test(six)) {
    throw new Error(`${field} is not a #RRGGBB color`);
  }
  return `#${six.toUpperCase()}`;
}

function mapLineStyle(raw: unknown): ChartLineStyle | string {
  if (typeof raw !== "string") throw new Error("ChartLineFormat.lineStyle is not a loaded string");
  const key = raw.toLowerCase().replace(/[^a-z0-9]/g, "");
  const table: Record<string, ChartLineStyle> = {
    none: "none",
    continuous: "continuous",
    dash: "dash",
    dashdot: "dashDot",
    dashdotdot: "dashDotDot",
    dot: "dot",
    grey25: "grey25",
    gray25: "grey25",
    grey50: "grey50",
    gray50: "grey50",
    grey75: "grey75",
    gray75: "grey75",
    automatic: "automatic",
    rounddot: "roundDot",
  };
  return table[key] ?? raw;
}

function toInfo(
  line: LineSurface,
  sheetName: string,
  chartName: string,
  seriesIndex: number,
  trendlineIndex: number,
): ChartTrendlineFormatInfo {
  return {
    sheetName,
    chartName,
    seriesIndex,
    trendlineIndex,
    color: normalizeTrendlineLineColor(
      requireLoadedString(line.color, "ChartLineFormat.color"),
      "ChartLineFormat.color",
    ),
    lineStyle: mapLineStyle(line.lineStyle),
    weight: requireLoadedNumber(line.weight, "ChartLineFormat.weight"),
  };
}

function assertMatches(
  info: ChartTrendlineFormatInfo,
  expected: { color?: string; lineStyle?: ChartLineStyle; weight?: number },
): void {
  if (expected.color !== undefined && info.color !== expected.color) {
    throw new Error(`color readback mismatch: expected ${expected.color}, got ${info.color}`);
  }
  if (expected.lineStyle !== undefined && info.lineStyle !== expected.lineStyle) {
    throw new Error(
      `lineStyle readback mismatch: expected ${expected.lineStyle}, got ${info.lineStyle}`,
    );
  }
  if (expected.weight !== undefined && info.weight !== expected.weight) {
    throw new Error(`weight readback mismatch: expected ${expected.weight}, got ${info.weight}`);
  }
}

/** Update trendline line format; ExcelApi 1.7; write → sync → load → sync. */
export async function officeJsUpdateChartSeriesTrendlineFormat(
  input: ChartTrendlineFormatUpdateInput,
): Promise<HostResult<ChartTrendlineFormatInfo>> {
  if (!isSetSupported(REQ_17)) {
    return unsupported(
      "chart.series.trendlines.format.update",
      "office-js",
      "ExcelApi 1.7 is not supported in this host (Office.context.requirements.isSetSupported)",
      EVIDENCE,
    );
  }

  return withExcel("chart.series.trendlines.format.update", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName) as unknown as {
      charts: {
        getItem(name: string): {
          name: string;
          load(p: string): void;
          series: {
            getItemAt(index: number): {
              trendlines: { getItem(index: number): TrendlineSurface };
            };
          };
        };
      };
    };
    const chart = sheet.charts.getItem(input.chartName);
    const trendline = chart.series
      .getItemAt(input.seriesIndex - 1)
      .trendlines.getItem(input.trendlineIndex - 1);
    const line = trendline.format.line;
    const color =
      input.color !== undefined
        ? normalizeTrendlineLineColor(input.color, "color")
        : undefined;
    if (color !== undefined) line.color = color;
    if (input.lineStyle !== undefined) line.lineStyle = STYLE_TO_HOST[input.lineStyle];
    if (input.weight !== undefined) line.weight = input.weight;
    await context.sync();
    chart.load("name");
    line.load("color,lineStyle,weight");
    await context.sync();
    const chartName = requireLoadedString(chart.name, "Chart.name");
    const info = toInfo(line, input.sheetName, chartName, input.seriesIndex, input.trendlineIndex);
    assertMatches(info, {
      color,
      lineStyle: input.lineStyle,
      weight: input.weight,
    });
    return info;
  });
}
