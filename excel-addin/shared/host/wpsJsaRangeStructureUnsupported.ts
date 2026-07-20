import { getSheet, requireWorkbook, type WpsRange } from "./wpsJsaRuntime";
import type {
  RangeAutofitInfo,
  RangeAutofitInput,
  RangeDeleteInput,
  RangeInsertInput,
  RangeMutationInfo,
} from "./rangeStructureTypes";
import type { HostResult } from "./types";
import { fail, ok, unsupported } from "./types";

const RANGE_STRUCTURE_EVIDENCE =
  "No in-repo WPS JSA Range.Insert/Delete contract";

const AUTOFIT_EVIDENCE =
  "Assumed Range.Columns/Rows.AutoFit + ColumnWidth/RowHeight (desktop ExcelTemplatePrintActionService COM parity; not in bridge contract; not device-verified)";

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
  _input: RangeInsertInput,
): Promise<HostResult<RangeMutationInfo>> {
  return unsupported(
    "range.insert",
    "wps-jsa",
    "range.insert is not verified for WPS JSA",
    RANGE_STRUCTURE_EVIDENCE,
  );
}

export async function wpsDeleteRange(
  _input: RangeDeleteInput,
): Promise<HostResult<RangeMutationInfo>> {
  return unsupported(
    "range.delete",
    "wps-jsa",
    "range.delete is not verified for WPS JSA",
    RANGE_STRUCTURE_EVIDENCE,
  );
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
      address: String(range.Address ?? input.address),
      direction: input.direction,
      columnWidth: requireNullableDimension(range.ColumnWidth, "Range.ColumnWidth"),
      rowHeight: requireNullableDimension(range.RowHeight, "Range.RowHeight"),
    });
  } catch (error) {
    return fail("range.autofit", "wps-jsa", messageOf(error), AUTOFIT_EVIDENCE);
  }
}
