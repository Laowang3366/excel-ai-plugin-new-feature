import { withExcel } from "./officeJsRuntime";
import type {
  ChartDataLabelPosition,
  ChartDataLabelsInfo,
  ChartDataLabelsUpdateInput,
} from "./chartDataLabelsTypes";
import type { HostResult } from "./types";
import { unsupported } from "./types";

const REQUIREMENT_SET = "ExcelApi";
const ENABLED_VERSION = "1.7";
const DATA_LABELS_VERSION = "1.8";

const POSITION_TO_HOST: Record<ChartDataLabelPosition, string> = {
  none: "None",
  center: "Center",
  insideEnd: "InsideEnd",
  insideBase: "InsideBase",
  outsideEnd: "OutsideEnd",
  left: "Left",
  right: "Right",
  top: "Top",
  bottom: "Bottom",
  bestFit: "BestFit",
  callout: "Callout",
};

interface ExcelDataLabels {
  showValue: boolean;
  showCategoryName: boolean;
  showSeriesName: boolean;
  numberFormat: string;
  showPercentage: boolean;
  showBubbleSize: boolean;
  showLegendKey: boolean;
  separator: string;
  position: string;
  load(props: string): void;
}

interface ExcelChartSeries {
  hasDataLabels: boolean;
  dataLabels: ExcelDataLabels;
  load(props: string): void;
}

interface ExcelChartSeriesCollection {
  getItemAt(index: number): ExcelChartSeries;
}

function requireLoadedString(value: unknown, field: string): string {
  if (typeof value !== "string") throw new Error(`${field} is not a loaded string`);
  return value;
}

function requireLoadedBoolean(value: unknown, field: string): boolean {
  if (typeof value !== "boolean") throw new Error(`${field} is not a loaded boolean`);
  return value;
}

function mapPositionFromHost(raw: unknown): ChartDataLabelPosition {
  if (typeof raw !== "string") {
    throw new Error("ChartDataLabels.position is not a loaded string");
  }
  const key = raw.toLowerCase().replace(/[^a-z]/g, "");
  const table: Record<string, ChartDataLabelPosition> = {
    none: "none",
    center: "center",
    insideend: "insideEnd",
    insidebase: "insideBase",
    outsideend: "outsideEnd",
    left: "left",
    right: "right",
    top: "top",
    bottom: "bottom",
    bestfit: "bestFit",
    callout: "callout",
  };
  const mapped = table[key];
  if (!mapped) {
    throw new Error(`ChartDataLabels.position has unsupported host value: ${raw}`);
  }
  return mapped;
}

function touchesDataLabelsFields(input: ChartDataLabelsUpdateInput): boolean {
  return (
    input.showValue !== undefined ||
    input.showCategoryName !== undefined ||
    input.showSeriesName !== undefined ||
    input.numberFormat !== undefined ||
    input.showPercentage !== undefined ||
    input.showBubbleSize !== undefined ||
    input.showLegendKey !== undefined ||
    input.separator !== undefined ||
    input.position !== undefined
  );
}

/** Official precheck before any dataLabels / hasDataLabels write. */
export function isExcelApiSupportedForDataLabels(version: "1.7" | "1.8"): boolean {
  const office = (globalThis as unknown as {
    Office?: {
      context?: {
        requirements?: { isSetSupported?: (name: string, minVersion?: string) => boolean };
      };
    };
  }).Office;
  const isSetSupported = office?.context?.requirements?.isSetSupported;
  if (typeof isSetSupported !== "function") return false;
  try {
    return isSetSupported.call(office!.context!.requirements, REQUIREMENT_SET, version);
  } catch {
    return false;
  }
}

const LOAD_PROPS =
  "showValue,showCategoryName,showSeriesName,numberFormat,showPercentage,showBubbleSize,showLegendKey,separator,position";

/**
 * Update series dataLabels/hasDataLabels.
 * - enabled-only: ExcelApi 1.7; never touches series.dataLabels; result omits show fields.
 * - any dataLabels field (optionally with enabled): ExcelApi 1.8 full snapshot path.
 */
export async function officeJsUpdateChartDataLabels(
  input: ChartDataLabelsUpdateInput,
): Promise<HostResult<ChartDataLabelsInfo>> {
  const needDataLabels = touchesDataLabelsFields(input);
  const version = needDataLabels ? DATA_LABELS_VERSION : ENABLED_VERSION;
  if (!isExcelApiSupportedForDataLabels(version)) {
    return unsupported(
      "chart.series.dataLabels.update",
      "office-js",
      `ExcelApi ${version} is not supported in this host (Office.context.requirements.isSetSupported)`,
      version === "1.8"
        ? "ChartSeries.dataLabels requires ExcelApi 1.8 (full snapshot path)"
        : "ChartSeries.hasDataLabels requires ExcelApi 1.7 (enabled-only path)",
    );
  }

  return withExcel("chart.series.dataLabels.update", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName) as unknown as {
      name: string;
      load(props: string): void;
      charts: {
        getItem(name: string): {
          name: string;
          load(props: string): void;
          series: ExcelChartSeriesCollection;
        };
      };
    };
    const chart = sheet.charts.getItem(input.chartName);
    const series = chart.series.getItemAt(input.seriesIndex - 1) as ExcelChartSeries & {
      hasDataLabels?: boolean;
      dataLabels?: ExcelDataLabels;
    };

    if (input.enabled !== undefined) {
      if (!("hasDataLabels" in series)) {
        throw new Error("hasDataLabels missing on ChartSeries");
      }
      series.hasDataLabels = input.enabled;
    }

    if (!needDataLabels) {
      await context.sync();
      sheet.load("name");
      chart.load("name");
      series.load("hasDataLabels");
      await context.sync();
      return {
        sheetName: requireLoadedString(sheet.name, "Worksheet.name"),
        chartName: requireLoadedString(chart.name, "Chart.name"),
        seriesIndex: input.seriesIndex,
        enabled: requireLoadedBoolean(series.hasDataLabels, "ChartSeries.hasDataLabels"),
      };
    }

    if (!("dataLabels" in series) || !series.dataLabels) {
      throw new Error("dataLabels missing on ChartSeries");
    }
    const labels = series.dataLabels;

    if (input.showValue !== undefined) labels.showValue = input.showValue;
    if (input.showCategoryName !== undefined) labels.showCategoryName = input.showCategoryName;
    if (input.showSeriesName !== undefined) labels.showSeriesName = input.showSeriesName;
    if (input.numberFormat !== undefined) labels.numberFormat = input.numberFormat;
    if (input.showPercentage !== undefined) labels.showPercentage = input.showPercentage;
    if (input.showBubbleSize !== undefined) labels.showBubbleSize = input.showBubbleSize;
    if (input.showLegendKey !== undefined) labels.showLegendKey = input.showLegendKey;
    if (input.separator !== undefined) labels.separator = input.separator;
    if (input.position !== undefined) {
      labels.position = POSITION_TO_HOST[input.position];
    }

    await context.sync();
    sheet.load("name");
    chart.load("name");
    series.load("hasDataLabels");
    labels.load(LOAD_PROPS);
    await context.sync();

    return {
      sheetName: requireLoadedString(sheet.name, "Worksheet.name"),
      chartName: requireLoadedString(chart.name, "Chart.name"),
      seriesIndex: input.seriesIndex,
      enabled: requireLoadedBoolean(series.hasDataLabels, "ChartSeries.hasDataLabels"),
      showValue: requireLoadedBoolean(labels.showValue, "ChartDataLabels.showValue"),
      showCategoryName: requireLoadedBoolean(
        labels.showCategoryName,
        "ChartDataLabels.showCategoryName",
      ),
      showSeriesName: requireLoadedBoolean(
        labels.showSeriesName,
        "ChartDataLabels.showSeriesName",
      ),
      numberFormat: requireLoadedString(labels.numberFormat, "ChartDataLabels.numberFormat"),
      showPercentage: requireLoadedBoolean(
        labels.showPercentage,
        "ChartDataLabels.showPercentage",
      ),
      showBubbleSize: requireLoadedBoolean(
        labels.showBubbleSize,
        "ChartDataLabels.showBubbleSize",
      ),
      showLegendKey: requireLoadedBoolean(labels.showLegendKey, "ChartDataLabels.showLegendKey"),
      separator: requireLoadedString(labels.separator, "ChartDataLabels.separator"),
      position: mapPositionFromHost(labels.position),
    };
  });
}
