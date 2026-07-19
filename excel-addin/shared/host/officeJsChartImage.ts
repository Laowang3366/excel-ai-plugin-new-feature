import { withExcel } from "./officeJsRuntime";
import type { ChartImageGetInput, ChartImageInfo } from "./chartImageTypes";
import type { HostResult } from "./types";
import { unsupported } from "./types";

const REQUIREMENT_EVIDENCE =
  "Chart.getImage requires ExcelApi 1.2 for host-generated Base64 image readback";

type ClientResult = { value: string };

function requireLoadedString(value: unknown, field: string): string {
  if (typeof value !== "string") throw new Error(`${field} is not a loaded string`);
  return value;
}

function requireImageBase64(value: unknown): string {
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error("Chart.getImage did not return a non-empty Base64 string");
  }
  return value;
}

/** Official precheck before any getImage call. */
export function isExcelApi12Supported(): boolean {
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
    return isSetSupported.call(office!.context!.requirements, "ExcelApi", "1.2");
  } catch {
    return false;
  }
}

/** Read chart image as Base64; precheck 1.2 → getImage → sync → host names + value. */
export async function officeJsGetChartImage(
  input: ChartImageGetInput,
): Promise<HostResult<ChartImageInfo>> {
  if (!isExcelApi12Supported()) {
    return unsupported(
      "chart.image.get",
      "office-js",
      "ExcelApi 1.2 is not supported in this host (Office.context.requirements.isSetSupported)",
      REQUIREMENT_EVIDENCE,
    );
  }

  const result = await withExcel("chart.image.get", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName) as unknown as {
      name: string;
      load(props: string): void;
      charts: {
        getItem(name: string): {
          name: string;
          load(props: string): void;
          getImage(width?: number, height?: number): ClientResult;
        };
      };
    };
    const chart = sheet.charts.getItem(input.chartName);
    if (typeof chart.getImage !== "function") {
      throw new Error("Chart.getImage missing (ExcelApi 1.2 required)");
    }
    sheet.load("name");
    chart.load("name");
    const imageResult =
      input.width != null || input.height != null
        ? chart.getImage(input.width, input.height)
        : chart.getImage();
    await context.sync();
    return {
      sheetName: requireLoadedString(sheet.name, "Worksheet.name"),
      chartName: requireLoadedString(chart.name, "Chart.name"),
      imageBase64: requireImageBase64(imageResult.value),
    };
  });
  if (
    !result.ok &&
    result.unsupported === true &&
    /getImage|ExcelApi 1\.2|requirement|isSetSupported/i.test(result.reason ?? "")
  ) {
    return unsupported(
      "chart.image.get",
      "office-js",
      result.reason ?? "ExcelApi 1.2 required for Chart.getImage",
      REQUIREMENT_EVIDENCE,
    );
  }
  return result;
}
