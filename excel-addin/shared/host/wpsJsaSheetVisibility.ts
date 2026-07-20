import { getSheet, requireWorkbook } from "./wpsJsaRuntime";
import type { HostResult, SheetVisibility, SheetVisibilityInfo } from "./types";
import { fail, ok, unsupported } from "./types";

/**
 * Assumed Worksheet.Visible (desktop ExcelObjectActionService COM parity).
 * Not in bridge contract; not device-verified.
 */
const EVIDENCE =
  "Assumed Worksheet.Visible xlSheetVisible=-1/Hidden=0/VeryHidden=2 (desktop COM parity; not in bridge contract; not device-verified)";

const VISIBLE = -1;
const HIDDEN = 0;
const VERY_HIDDEN = 2;

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export function mapVisibleToPublic(raw: unknown): SheetVisibility | null {
  if (raw === VISIBLE || raw === true) return "visible";
  if (raw === HIDDEN || raw === false) return "hidden";
  if (raw === VERY_HIDDEN) return "veryHidden";
  const text = String(raw ?? "")
    .toLowerCase()
    .replace(/\s+/g, "");
  if (text.includes("veryhidden") || text === "2") return "veryHidden";
  if (text.includes("hidden") || text === "0") return "hidden";
  if (text.includes("visible") || text === "-1" || text === "1") return "visible";
  return null;
}

export function mapPublicToVisible(v: SheetVisibility): number {
  if (v === "veryHidden") return VERY_HIDDEN;
  if (v === "hidden") return HIDDEN;
  return VISIBLE;
}

function isSheetVisibility(value: unknown): value is SheetVisibility {
  return value === "visible" || value === "hidden" || value === "veryHidden";
}

export async function wpsGetSheetVisibility(
  sheetName: string,
): Promise<HostResult<SheetVisibilityInfo>> {
  const workbookResult = requireWorkbook("sheet.visibility.get");
  if (!workbookResult.ok) return workbookResult;
  const sheet = getSheet(workbookResult.data, sheetName);
  if (!sheet) {
    return unsupported(
      "sheet.visibility.get",
      "wps-jsa",
      `Sheet "${sheetName}" not found`,
      EVIDENCE,
    );
  }
  if (!("Visible" in sheet) || sheet.Visible === undefined) {
    return unsupported(
      "sheet.visibility.get",
      "wps-jsa",
      "Worksheet.Visible is unavailable",
      EVIDENCE,
    );
  }
  try {
    const mapped = mapVisibleToPublic(sheet.Visible);
    if (!mapped) {
      return fail(
        "sheet.visibility.get",
        "wps-jsa",
        `unrecognized Visible value: ${String(sheet.Visible)}`,
        EVIDENCE,
      );
    }
    return ok({ sheetName: sheet.Name, visibility: mapped });
  } catch (error) {
    return fail("sheet.visibility.get", "wps-jsa", messageOf(error), EVIDENCE);
  }
}

export async function wpsSetSheetVisibility(
  sheetName: string,
  visibility: SheetVisibility | string,
): Promise<HostResult<SheetVisibilityInfo>> {
  if (!isSheetVisibility(visibility)) {
    return fail(
      "sheet.visibility.set",
      "wps-jsa",
      `visibility must be visible|hidden|veryHidden, got "${String(visibility)}"`,
      EVIDENCE,
    );
  }
  const workbookResult = requireWorkbook("sheet.visibility.set");
  if (!workbookResult.ok) return workbookResult;
  const sheet = getSheet(workbookResult.data, sheetName);
  if (!sheet) {
    return unsupported(
      "sheet.visibility.set",
      "wps-jsa",
      `Sheet "${sheetName}" not found`,
      EVIDENCE,
    );
  }
  if (!("Visible" in sheet) || sheet.Visible === undefined) {
    return unsupported(
      "sheet.visibility.set",
      "wps-jsa",
      "Worksheet.Visible is unavailable",
      EVIDENCE,
    );
  }
  try {
    sheet.Visible = mapPublicToVisible(visibility);
    const mapped = mapVisibleToPublic(sheet.Visible);
    if (!mapped) {
      return fail(
        "sheet.visibility.set",
        "wps-jsa",
        `Visible writeback unrecognized: ${String(sheet.Visible)}`,
        EVIDENCE,
      );
    }
    if (mapped !== visibility) {
      return fail(
        "sheet.visibility.set",
        "wps-jsa",
        `Visible writeback mismatch: wanted ${visibility}, got ${mapped}`,
        EVIDENCE,
      );
    }
    return ok({ sheetName: sheet.Name, visibility: mapped });
  } catch (error) {
    return fail("sheet.visibility.set", "wps-jsa", messageOf(error), EVIDENCE);
  }
}
