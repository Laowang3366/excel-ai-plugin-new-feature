import { withExcel } from "./officeJsRuntime";
import type {
  ChartDataLabelsInfo,
  ChartDataLabelsUpdateInput,
} from "./chartDataLabelsTypes";
import type { HostResult } from "./types";
import { unsupported } from "./types";

const REQUIREMENT_SET = "ExcelApi";
const ENABLED_VERSION = "1.7";
const DATA_LABELS_VERSION = "1.8";

interface ExcelDataLabels {
  showValue: boolean;
  showCategoryName: boolean;
  showSeriesName: boolean;
  numberFormat: string;
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

/** Only explicit requirement-set precheck failures — not missing members / business errors. */

function touchesDataLabelsFields(input: ChartDataLabelsUpdateInput): boolean {
  return (
    input.showValue !== undefined ||
    input.showCategoryName !== undefined ||
    input.showSeriesName !== undefined ||
    input.numberFormat !== undefined
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

/**
 * Update series dataLabels/hasDataLabels.
 * - enabled-only: ExcelApi 1.7; never touches series.dataLabels; result omits show fields.
 * - any show field or numberFormat (optionally with enabled): ExcelApi 1.8 full snapshot path.
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

  const result = await withExcel("chart.series.dataLabels.update", async (context) => {
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
      // ExcelApi 1.7 enabled-only: do not access series.dataLabels.
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

    await context.sync();
    sheet.load("name");
    chart.load("name");
    series.load("hasDataLabels");
    labels.load("showValue,showCategoryName,showSeriesName,numberFormat");
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
    };
  });

  return result;
}
