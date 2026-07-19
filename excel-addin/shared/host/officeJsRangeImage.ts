import { getExcelRun } from "./officeJsRuntime";
import type { RangeImageGetInput, RangeImageInfo } from "./rangeImageTypes";
import type { HostResult } from "./types";
import { fail, ok, unsupported } from "./types";

const REQUIREMENT_EVIDENCE =
  "Range.getImage requires ExcelApi 1.7 for host-generated Base64 PNG readback";
const EXCEL_RUN_EVIDENCE = "Requires Office.js Excel.run";
const CAPABILITY = "range.image.get";

type ClientResult = { value: string };

function requireLoadedString(value: unknown, field: string): string {
  if (typeof value !== "string") throw new Error(`${field} is not a loaded string`);
  return value;
}

function requireImageBase64(value: unknown): string {
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error("Range.getImage did not return a non-empty Base64 string");
  }
  return value;
}

/** Official precheck before any getImage call. */
export function isExcelApi17Supported(): boolean {
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
    return isSetSupported.call(office!.context!.requirements, "ExcelApi", "1.7");
  } catch {
    return false;
  }
}

/**
 * Local runner: missing Excel.run is typed unsupported; after precheck, batch errors are ordinary fail.
 * Does not use withExcel (which marks all catch as unsupported).
 */
async function runRangeImage(
  fn: (context: {
    workbook: {
      worksheets: {
        getItem(name: string): {
          name: string;
          load(props: string): void;
          getRange(address: string): {
            address: string;
            load(props: string): void;
            getImage(): ClientResult;
          };
        };
      };
    };
    sync(): Promise<void>;
  }) => Promise<RangeImageInfo>,
): Promise<HostResult<RangeImageInfo>> {
  const run = getExcelRun();
  if (!run) {
    return unsupported(
      CAPABILITY,
      "office-js",
      "Excel.run is not available in this runtime",
      EXCEL_RUN_EVIDENCE,
    );
  }
  try {
    return ok(await run(fn as never));
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return fail(CAPABILITY, "office-js", message);
  }
}

/** Read range image as Base64; precheck 1.7 → getImage → sync → host name/address/value. */
export async function officeJsGetRangeImage(
  input: RangeImageGetInput,
): Promise<HostResult<RangeImageInfo>> {
  if (!isExcelApi17Supported()) {
    return unsupported(
      CAPABILITY,
      "office-js",
      "ExcelApi 1.7 is not supported in this host (Office.context.requirements.isSetSupported)",
      REQUIREMENT_EVIDENCE,
    );
  }

  return runRangeImage(async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    const range = sheet.getRange(input.range);
    if (typeof range.getImage !== "function") {
      throw new Error("Range.getImage missing (ExcelApi 1.7 required)");
    }
    sheet.load("name");
    range.load("address");
    const imageResult = range.getImage();
    await context.sync();
    return {
      sheetName: requireLoadedString(sheet.name, "Worksheet.name"),
      address: requireLoadedString(range.address, "Range.address"),
      imageBase64: requireImageBase64(imageResult.value),
    };
  });
}
