import type {
  RangeAutofitInfo,
  RangeAutofitInput,
  RangeDeleteInput,
  RangeInsertInput,
  RangeMutationInfo,
} from "./rangeStructureTypes";
import { withExcel } from "./officeJsRuntime";
import type { HostResult } from "./types";
import { unsupported } from "./types";

const INSERT_SHIFT = { down: "Down", right: "Right" } as const;
const DELETE_SHIFT = { up: "Up", left: "Left" } as const;

function requireLoadedAddress(value: unknown): string {
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error("Range.address is not a loaded non-empty string");
  }
  return value;
}

function requireNullableDimension(value: unknown, field: string): number | null {
  if (value === null) return null;
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error(`${field} is not a loaded finite number or null`);
  }
  return value;
}

function isExcelApi12Supported(): boolean {
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

export async function officeJsInsertRange(
  input: RangeInsertInput,
): Promise<HostResult<RangeMutationInfo>> {
  return withExcel("range.insert", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    const inserted = sheet.getRange(input.address).insert(INSERT_SHIFT[input.shift]);
    await context.sync();
    inserted.load("address");
    await context.sync();
    return {
      sheetName: input.sheetName,
      address: requireLoadedAddress(inserted.address),
      shift: input.shift,
      operation: "insert",
    };
  });
}

export async function officeJsDeleteRange(
  input: RangeDeleteInput,
): Promise<HostResult<RangeMutationInfo>> {
  return withExcel("range.delete", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    const range = sheet.getRange(input.address);
    range.load("address");
    await context.sync();
    const address = requireLoadedAddress(range.address);
    range.delete(DELETE_SHIFT[input.shift]);
    await context.sync();
    return {
      sheetName: input.sheetName,
      address,
      shift: input.shift,
      operation: "delete",
    };
  });
}

export async function officeJsAutofitRange(
  input: RangeAutofitInput,
): Promise<HostResult<RangeAutofitInfo>> {
  if (!isExcelApi12Supported()) {
    return unsupported(
      "range.autofit",
      "office-js",
      "ExcelApi 1.2 is not supported in this host (Office.context.requirements.isSetSupported)",
      "RangeFormat.autofitColumns/autofitRows and dimension readback require ExcelApi 1.2",
    );
  }
  return withExcel("range.autofit", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    const range = sheet.getRange(input.address);
    if (input.direction !== "rows") range.format.autofitColumns();
    if (input.direction !== "columns") range.format.autofitRows();
    await context.sync();
    range.load("address");
    range.format.load("columnWidth,rowHeight");
    await context.sync();
    return {
      sheetName: input.sheetName,
      address: requireLoadedAddress(range.address),
      direction: input.direction,
      columnWidth: requireNullableDimension(range.format.columnWidth, "RangeFormat.columnWidth"),
      rowHeight: requireNullableDimension(range.format.rowHeight, "RangeFormat.rowHeight"),
    };
  });
}
