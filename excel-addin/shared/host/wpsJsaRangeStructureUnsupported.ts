import { getSheet, requireWorkbook, type WpsRange } from "./wpsJsaRuntime";
import { readWpsAddress } from "./wpsJsaAddress";
import type {
  RangeAutofitInfo,
  RangeAutofitInput,
  RangeDeleteInput,
  RangeInsertInput,
  RangeMutationInfo,
} from "./rangeStructureTypes";
import type { HostResult } from "./types";
import { fail, ok, unsupported } from "./types";

const AUTOFIT_EVIDENCE =
  "Assumed Range.Columns/Rows.AutoFit + ColumnWidth/RowHeight (desktop ExcelTemplatePrintActionService COM parity; not in bridge contract; not device-verified)";

/**
 * Excel COM Shift constants (xlShift*). Used as host contract for Insert/Delete.
 * Not in JSA bridge; member-probed at runtime.
 */
const XL_SHIFT_DOWN = -4121;
const XL_SHIFT_TO_RIGHT = -4161;
const XL_SHIFT_UP = -4162;
const XL_SHIFT_TO_LEFT = -4159;

const INSERT_SHIFT = {
  down: XL_SHIFT_DOWN,
  right: XL_SHIFT_TO_RIGHT,
} as const;

const DELETE_SHIFT = {
  up: XL_SHIFT_UP,
  left: XL_SHIFT_TO_LEFT,
} as const;

const INSERT_DELETE_EVIDENCE =
  "Assumed Range.Insert/Delete with xlShift constants (ET COM parity; not in bridge contract; not device-verified)";

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function requireNullableDimension(value: unknown, field: string): number | null {
  if (value === null || value === undefined) return null;
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error(`${field} is not a finite number or null`);
  }
  return value;
}

function pickAutoFitTarget(
  range: WpsRange,
  kind: "columns" | "rows",
): (() => void) | null {
  if (kind === "columns") {
    if (typeof range.Columns?.AutoFit === "function") {
      return () => range.Columns!.AutoFit!();
    }
    if (typeof range.EntireColumn?.AutoFit === "function") {
      return () => range.EntireColumn!.AutoFit!();
    }
    return null;
  }
  if (typeof range.Rows?.AutoFit === "function") {
    return () => range.Rows!.AutoFit!();
  }
  if (typeof range.EntireRow?.AutoFit === "function") {
    return () => range.EntireRow!.AutoFit!();
  }
  return null;
}

export async function wpsInsertRange(
  input: RangeInsertInput,
): Promise<HostResult<RangeMutationInfo>> {
  if (input.shift !== "down" && input.shift !== "right") {
    return fail(
      "range.insert",
      "wps-jsa",
      `shift must be down|right, got "${String(input.shift)}"`,
      INSERT_DELETE_EVIDENCE,
    );
  }
  const workbookResult = requireWorkbook("range.insert");
  if (!workbookResult.ok) return workbookResult;
  const sheet = getSheet(workbookResult.data, input.sheetName);
  if (!sheet?.Range) {
    return unsupported(
      "range.insert",
      "wps-jsa",
      `Sheet "${input.sheetName}" or Range API missing`,
      INSERT_DELETE_EVIDENCE,
    );
  }
  let range: WpsRange;
  try {
    range = sheet.Range(input.address);
  } catch (error) {
    return fail("range.insert", "wps-jsa", messageOf(error), INSERT_DELETE_EVIDENCE);
  }
  if (typeof range.Insert !== "function") {
    return unsupported(
      "range.insert",
      "wps-jsa",
      "Range.Insert is unavailable",
      INSERT_DELETE_EVIDENCE,
    );
  }
  try {
    const hostShift = INSERT_SHIFT[input.shift];
    const inserted = range.Insert(hostShift);
    const resultRange =
      inserted && typeof inserted === "object" ? (inserted as WpsRange) : range;
    const address =
      readWpsAddress(resultRange, readWpsAddress(range, input.address) ?? input.address) ??
      input.address;
    if (!address) {
      return fail(
        "range.insert",
        "wps-jsa",
        "Insert completed but Address is unavailable",
        INSERT_DELETE_EVIDENCE,
      );
    }
    return ok({
      sheetName: input.sheetName,
      address,
      shift: input.shift,
      operation: "insert",
    });
  } catch (error) {
    return fail("range.insert", "wps-jsa", messageOf(error), INSERT_DELETE_EVIDENCE);
  }
}

export async function wpsDeleteRange(
  input: RangeDeleteInput,
): Promise<HostResult<RangeMutationInfo>> {
  if (input.shift !== "up" && input.shift !== "left") {
    return fail(
      "range.delete",
      "wps-jsa",
      `shift must be up|left, got "${String(input.shift)}"`,
      INSERT_DELETE_EVIDENCE,
    );
  }
  const workbookResult = requireWorkbook("range.delete");
  if (!workbookResult.ok) return workbookResult;
  const sheet = getSheet(workbookResult.data, input.sheetName);
  if (!sheet?.Range) {
    return unsupported(
      "range.delete",
      "wps-jsa",
      `Sheet "${input.sheetName}" or Range API missing`,
      INSERT_DELETE_EVIDENCE,
    );
  }
  let range: WpsRange;
  try {
    range = sheet.Range(input.address);
  } catch (error) {
    return fail("range.delete", "wps-jsa", messageOf(error), INSERT_DELETE_EVIDENCE);
  }
  if (typeof range.Delete !== "function") {
    return unsupported(
      "range.delete",
      "wps-jsa",
      "Range.Delete is unavailable",
      INSERT_DELETE_EVIDENCE,
    );
  }
  try {
    const address = readWpsAddress(range, input.address) ?? input.address;
    if (!address) {
      return fail(
        "range.delete",
        "wps-jsa",
        "Range.Address is unavailable before Delete",
        INSERT_DELETE_EVIDENCE,
      );
    }
    range.Delete(DELETE_SHIFT[input.shift]);
    return ok({
      sheetName: input.sheetName,
      address,
      shift: input.shift,
      operation: "delete",
    });
  } catch (error) {
    return fail("range.delete", "wps-jsa", messageOf(error), INSERT_DELETE_EVIDENCE);
  }
}

export async function wpsAutofitRange(
  input: RangeAutofitInput,
): Promise<HostResult<RangeAutofitInfo>> {
  const workbookResult = requireWorkbook("range.autofit");
  if (!workbookResult.ok) return workbookResult;
  const sheet = getSheet(workbookResult.data, input.sheetName);
  if (!sheet?.Range) {
    return unsupported(
      "range.autofit",
      "wps-jsa",
      `Sheet "${input.sheetName}" or Range API missing`,
      AUTOFIT_EVIDENCE,
    );
  }

  let range: WpsRange;
  try {
    range = sheet.Range(input.address);
  } catch (error) {
    return fail("range.autofit", "wps-jsa", messageOf(error), AUTOFIT_EVIDENCE);
  }

  const runColumns = input.direction !== "rows";
  const runRows = input.direction !== "columns";

  if (runColumns) {
    const fit = pickAutoFitTarget(range, "columns");
    if (!fit) {
      return unsupported(
        "range.autofit",
        "wps-jsa",
        "Range.Columns/EntireColumn.AutoFit is unavailable",
        AUTOFIT_EVIDENCE,
      );
    }
    try {
      fit();
    } catch (error) {
      return fail("range.autofit", "wps-jsa", messageOf(error), AUTOFIT_EVIDENCE);
    }
  }

  if (runRows) {
    const fit = pickAutoFitTarget(range, "rows");
    if (!fit) {
      return unsupported(
        "range.autofit",
        "wps-jsa",
        "Range.Rows/EntireRow.AutoFit is unavailable",
        AUTOFIT_EVIDENCE,
      );
    }
    try {
      fit();
    } catch (error) {
      return fail("range.autofit", "wps-jsa", messageOf(error), AUTOFIT_EVIDENCE);
    }
  }

  try {
    return ok({
      sheetName: input.sheetName,
      address: readWpsAddress(range, input.address) ?? input.address,
      direction: input.direction,
      columnWidth: requireNullableDimension(range.ColumnWidth, "Range.ColumnWidth"),
      rowHeight: requireNullableDimension(range.RowHeight, "Range.RowHeight"),
    });
  } catch (error) {
    return fail("range.autofit", "wps-jsa", messageOf(error), AUTOFIT_EVIDENCE);
  }
}
