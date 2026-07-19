import { normalizeSameSheetSourceRange } from "./officeJsChartSource";
import { withExcel } from "./officeJsRuntime";
import type {
  ChartSeriesValuesInfo,
  ChartSeriesValuesUpdateInput,
} from "./chartSeriesValuesTypes";
import type { HostResult } from "./types";
import { unsupported } from "./types";

const REQUIREMENT_SET = "ExcelApi";
const REQUIREMENT_VERSION = "1.15";
const REQUIREMENT_EVIDENCE =
  "getDimensionDataSourceString requires ExcelApi 1.15 for verified series source readback";

type ClientResult = { value: string };

type SeriesSurface = {
  setValues(range: object): void;
  setXAxisValues(range: object): void;
  getDimensionDataSourceString?(dimension: "Values" | "XValues"): ClientResult;
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


/** Official precheck before any series data-source write. */
export function isExcelApi115Supported(): boolean {
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

/** Bind series values/xValues; precheck 1.15 → write → sync → getDimensionDataSourceString → sync. */
export async function officeJsUpdateChartSeriesValues(
  input: ChartSeriesValuesUpdateInput,
): Promise<HostResult<ChartSeriesValuesInfo>> {
  if (!isExcelApi115Supported()) {
    return unsupported(
      "chart.series.values.update",
      "office-js",
      "ExcelApi 1.15 is not supported in this host (Office.context.requirements.isSetSupported)",
      REQUIREMENT_EVIDENCE,
    );
  }

  const result = await withExcel("chart.series.values.update", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    const chart = sheet.charts.getItem(input.chartName) as unknown as {
      name: string;
      load(props: string): void;
      series: { getItemAt(index: number): SeriesSurface };
    };
    const series = chart.series.getItemAt(input.seriesIndex - 1);
    if (typeof series.getDimensionDataSourceString !== "function") {
      throw new Error("getDimensionDataSourceString missing (ExcelApi 1.15 required)");
    }
    if (input.valuesRange != null) {
      const bare = normalizeSameSheetSourceRange(input.sheetName, input.valuesRange);
      series.setValues(sheet.getRange(bare));
    }
    if (input.xValuesRange != null) {
      const bare = normalizeSameSheetSourceRange(input.sheetName, input.xValuesRange);
      series.setXAxisValues(sheet.getRange(bare));
    }
    await context.sync();
    chart.load("name");
    const valuesResult =
      input.valuesRange != null
        ? series.getDimensionDataSourceString!("Values")
        : undefined;
    const xValuesResult =
      input.xValuesRange != null
        ? series.getDimensionDataSourceString!("XValues")
        : undefined;
    await context.sync();
    const info: ChartSeriesValuesInfo = {
      sheetName: input.sheetName,
      chartName: requireLoadedString(chart.name, "Chart.name"),
      seriesIndex: input.seriesIndex,
      dataBound: true,
    };
    if (valuesResult) info.valuesSource = readSource(valuesResult, "valuesSource");
    if (xValuesResult) info.xValuesSource = readSource(xValuesResult, "xValuesSource");
    return info;
  });
  return result;
}
