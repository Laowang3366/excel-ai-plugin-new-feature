import { getSheet, requireWorkbook } from "./wpsJsaRuntime";
import type { HostResult, SheetProtectionInfo } from "./types";
import { fail, ok, unsupported } from "./types";

/**
 * Assumed Worksheet.ProtectContents / Protect / Unprotect
 * (desktop ExcelObjectActionService COM parity). Not device-verified.
 * Password stays request-memory only — never appear in result/reason/evidence.
 */
const EVIDENCE =
  "Assumed Worksheet.ProtectContents/Protect/Unprotect (desktop COM parity; not in bridge contract; not device-verified)";

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

/** Strip accidental password echoes from host error text. */
export function sanitizeProtectionMessage(message: string, password?: string): string {
  let out = message;
  if (password && password.length > 0) {
    out = out.split(password).join("[redacted]");
  }
  return out;
}

function readProtected(sheet: { ProtectContents?: boolean }): boolean | null {
  if (!("ProtectContents" in sheet) || sheet.ProtectContents === undefined) return null;
  return sheet.ProtectContents === true;
}

export async function wpsGetSheetProtection(
  sheetName: string,
): Promise<HostResult<SheetProtectionInfo>> {
  const workbookResult = requireWorkbook("sheet.protection.get");
  if (!workbookResult.ok) return workbookResult;
  const sheet = getSheet(workbookResult.data, sheetName);
  if (!sheet) {
    return unsupported(
      "sheet.protection.get",
      "wps-jsa",
      `Sheet "${sheetName}" not found`,
      EVIDENCE,
    );
  }
  const protectedState = readProtected(sheet);
  if (protectedState == null) {
    return unsupported(
      "sheet.protection.get",
      "wps-jsa",
      "Worksheet.ProtectContents is unavailable",
      EVIDENCE,
    );
  }
  try {
    return ok({ sheetName: sheet.Name, protected: protectedState });
  } catch (error) {
    return fail(
      "sheet.protection.get",
      "wps-jsa",
      sanitizeProtectionMessage(messageOf(error)),
      EVIDENCE,
    );
  }
}

export async function wpsProtectSheet(
  sheetName: string,
  password?: string,
): Promise<HostResult<SheetProtectionInfo>> {
  const workbookResult = requireWorkbook("sheet.protection.protect");
  if (!workbookResult.ok) return workbookResult;
  const sheet = getSheet(workbookResult.data, sheetName);
  if (!sheet) {
    return unsupported(
      "sheet.protection.protect",
      "wps-jsa",
      `Sheet "${sheetName}" not found`,
      EVIDENCE,
    );
  }
  if (typeof sheet.Protect !== "function") {
    return unsupported(
      "sheet.protection.protect",
      "wps-jsa",
      "Worksheet.Protect is unavailable",
      EVIDENCE,
    );
  }
  const before = readProtected(sheet);
  if (before === true) {
    return fail(
      "sheet.protection.protect",
      "wps-jsa",
      "sheet is already protected",
      EVIDENCE,
    );
  }
  try {
    if (password != null && password !== "") {
      sheet.Protect(password);
    } else {
      sheet.Protect();
    }
    const after = readProtected(sheet);
    if (after == null) {
      return unsupported(
        "sheet.protection.protect",
        "wps-jsa",
        "Protect called but ProtectContents is unavailable for verification",
        EVIDENCE,
      );
    }
    if (after !== true) {
      return fail(
        "sheet.protection.protect",
        "wps-jsa",
        "Protect completed but ProtectContents is still false",
        EVIDENCE,
      );
    }
    return ok({ sheetName: sheet.Name, protected: true });
  } catch (error) {
    return fail(
      "sheet.protection.protect",
      "wps-jsa",
      sanitizeProtectionMessage(messageOf(error), password),
      EVIDENCE,
    );
  }
}

export async function wpsUnprotectSheet(
  sheetName: string,
  password?: string,
): Promise<HostResult<SheetProtectionInfo>> {
  const workbookResult = requireWorkbook("sheet.protection.unprotect");
  if (!workbookResult.ok) return workbookResult;
  const sheet = getSheet(workbookResult.data, sheetName);
  if (!sheet) {
    return unsupported(
      "sheet.protection.unprotect",
      "wps-jsa",
      `Sheet "${sheetName}" not found`,
      EVIDENCE,
    );
  }
  if (typeof sheet.Unprotect !== "function") {
    return unsupported(
      "sheet.protection.unprotect",
      "wps-jsa",
      "Worksheet.Unprotect is unavailable",
      EVIDENCE,
    );
  }
  const before = readProtected(sheet);
  if (before === false) {
    return fail(
      "sheet.protection.unprotect",
      "wps-jsa",
      "sheet is not protected",
      EVIDENCE,
    );
  }
  try {
    if (password != null && password !== "") {
      sheet.Unprotect(password);
    } else {
      sheet.Unprotect();
    }
    const after = readProtected(sheet);
    if (after == null) {
      return unsupported(
        "sheet.protection.unprotect",
        "wps-jsa",
        "Unprotect called but ProtectContents is unavailable for verification",
        EVIDENCE,
      );
    }
    if (after !== false) {
      return fail(
        "sheet.protection.unprotect",
        "wps-jsa",
        "Unprotect completed but ProtectContents is still true",
        EVIDENCE,
      );
    }
    return ok({ sheetName: sheet.Name, protected: false });
  } catch (error) {
    return fail(
      "sheet.protection.unprotect",
      "wps-jsa",
      sanitizeProtectionMessage(messageOf(error), password),
      EVIDENCE,
    );
  }
}
