import { normalizeSameSheetSourceRange } from "./officeJsChartSource";
import { withExcel } from "./officeJsRuntime";
import type {
  ChartSeriesBubbleSizesInfo,
  ChartSeriesBubbleSizesUpdateInput,
} from "./chartSeriesBubbleSizesTypes";
import type { HostResult } from "./types";
import { unsupported } from "./types";

const REQUIREMENT_SET = "ExcelApi";
const REQUIREMENT_VERSION = "1.15";
const REQUIREMENT_EVIDENCE =
  "getDimensionDataSourceString(BubbleSizes) requires ExcelApi 1.15 for verified bubble size source readback";

type ClientResult = { value: string };

type SeriesSurface = {
  setBubbleSizes(range: object): void;
  getDimensionDataSourceString?(dimension: "BubbleSizes"): ClientResult;
};

function requireLoadedString(value: unknown, field: string): string {
  if (typeof value !== "string") throw new Error(`${field} is not a loaded string`);
  return value;
}

function readSource(result: ClientResult, field: string): string {
  const value = result.value;
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`${field} is not a loaded non-empty source string`);
  }
  return value;
}

/** Only true requirement-set signals — not setBubbleSizes/BubbleSizes business errors. */

/** Official precheck before any bubble-sizes write. */
export function isExcelApi115SupportedForBubbleSizes(): boolean {
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
    return isSetSupported.call(office!.context!.requirements, REQUIREMENT_SET, REQUIREMENT_VERSION);
  } catch {
    return false;
  }
}

/** Bind series bubble sizes; precheck 1.15 → setBubbleSizes → sync → readback → sync. */
export async function officeJsUpdateChartSeriesBubbleSizes(
  input: ChartSeriesBubbleSizesUpdateInput,
): Promise<HostResult<ChartSeriesBubbleSizesInfo>> {
  if (!isExcelApi115SupportedForBubbleSizes()) {
    return unsupported(
      "chart.series.bubbleSizes.update",
      "office-js",
      "ExcelApi 1.15 is not supported in this host (Office.context.requirements.isSetSupported)",
      REQUIREMENT_EVIDENCE,
    );
  }

  const result = await withExcel("chart.series.bubbleSizes.update", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName) as unknown as {
      name: string;
      load(props: string): void;
      getRange(address: string): object;
      charts: {
        getItem(name: string): {
          name: string;
          load(props: string): void;
          series: { getItemAt(index: number): SeriesSurface };
        };
      };
    };
    const chart = sheet.charts.getItem(input.chartName);
    const series = chart.series.getItemAt(input.seriesIndex - 1);
    if (typeof series.setBubbleSizes !== "function") {
      throw new Error("setBubbleSizes missing (ExcelApi 1.7 required)");
    }
    if (typeof series.getDimensionDataSourceString !== "function") {
      throw new Error("getDimensionDataSourceString missing (ExcelApi 1.15 required)");
    }
    const bare = normalizeSameSheetSourceRange(input.sheetName, input.bubbleSizesRange);
    series.setBubbleSizes(sheet.getRange(bare));
    await context.sync();
    sheet.load("name");
    chart.load("name");
    const sourceResult = series.getDimensionDataSourceString!("BubbleSizes");
    await context.sync();
    return {
      sheetName: requireLoadedString(sheet.name, "Worksheet.name"),
      chartName: requireLoadedString(chart.name, "Chart.name"),
      seriesIndex: input.seriesIndex,
      bubbleSizesSource: readSource(sourceResult, "bubbleSizesSource"),
      dataBound: true as const,
    };
  });
  return result;
}
