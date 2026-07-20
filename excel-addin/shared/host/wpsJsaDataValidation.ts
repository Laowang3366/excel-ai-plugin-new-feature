/**
 * WPS JSA data validation via Range.Validation (COM-style).
 * Covers protocol types/operators, allowBlank, inline list, A1 range source, clear.
 * Write snapshots prior validation and restores on Add failure when possible.
 */
import {
  classifyListSource,
  dvRulesMatch,
} from "./officeJsValidationCompare";
import {
  isBetweenOp,
  MAX_INLINE_LIST_SOURCE_CHARS,
} from "./officeJsValidationMapping";
import type {
  DataValidationInfo,
  DataValidationRule,
  HostResult,
} from "./types";
import { fail, ok, unsupported } from "./types";
import { getSheet, requireWorkbook, type WpsRange, type WpsValidation } from "./wpsJsaRuntime";
import {
  parseHostValidation,
  restoreSnapshot,
  trySnapshot,
  type ParsedWpsValidation,
} from "./wpsJsaDataValidationParse";
import {
  DV_EVIDENCE as EVIDENCE,
  XL_VALID_ALERT_STOP,
  XL_VALIDATE_CUSTOM,
  XL_VALIDATE_LIST,
  mapDvOperatorToCom,
  mapDvTypeToCom,
} from "./wpsJsaValidationConstants";

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function resolveRange(
  capability: string,
  sheetName: string,
  address: string,
): HostResult<{ range: WpsRange; sheetName: string; address: string }> {
  const workbookResult = requireWorkbook(capability);
  if (!workbookResult.ok) return workbookResult;
  const sheet = getSheet(workbookResult.data, sheetName);
  if (!sheet?.Range) {
    return unsupported(
      capability,
      "wps-jsa",
      `Sheet "${sheetName}" or Range API missing`,
      EVIDENCE,
    );
  }
  try {
    const range = sheet.Range(address);
    return ok({
      range,
      sheetName,
      address: String(range.Address ?? `${sheetName}!${address}`),
    });
  } catch (error) {
    return fail(capability, "wps-jsa", messageOf(error), EVIDENCE);
  }
}

function requireValidation(
  capability: string,
  range: WpsRange,
): HostResult<WpsValidation> {
  const v = range.Validation;
  if (!v || typeof v !== "object") {
    return unsupported(capability, "wps-jsa", "Range.Validation is unavailable", EVIDENCE);
  }
  return ok(v);
}

function toInfo(
  sheetName: string,
  rangeAddress: string,
  parsed: ParsedWpsValidation,
): DataValidationInfo {
  return {
    sheetName,
    range: rangeAddress,
    rule: parsed.rule,
    hostType: parsed.hostType,
    supported: parsed.supported,
    listSourceKind: parsed.listSourceKind ?? null,
    limitations: parsed.limitations,
  };
}

function buildListSource(rule: DataValidationRule): HostResult<string> {
  if (rule.listValues && rule.listValues.length > 0) {
    if (rule.listValues.some((v) => v.includes(","))) {
      return fail(
        "dataValidation.write",
        "wps-jsa",
        "listValues items must not contain commas; use a range source instead",
        EVIDENCE,
      );
    }
    const source = rule.listValues.join(",");
    if (source.length > MAX_INLINE_LIST_SOURCE_CHARS) {
      return fail(
        "dataValidation.write",
        "wps-jsa",
        `inline list source exceeds Excel ${MAX_INLINE_LIST_SOURCE_CHARS} character limit; use a range source`,
        EVIDENCE,
      );
    }
    return ok(source);
  }
  if (rule.formula1) {
    const classified = classifyListSource(rule.formula1);
    if (classified.kind !== "range" || classified.lossy) {
      return fail(
        "dataValidation.write",
        "wps-jsa",
        `list formula1 must be a same-workbook A1 range: ${rule.formula1}`,
        EVIDENCE,
      );
    }
    const body = classified.formula1 ?? rule.formula1;
    return ok(body.startsWith("=") ? body : `=${body}`);
  }
  return fail(
    "dataValidation.write",
    "wps-jsa",
    "list validation requires listValues or formula1 range source",
    EVIDENCE,
  );
}

export async function wpsReadDataValidation(
  sheetName: string,
  rangeAddress: string,
): Promise<HostResult<DataValidationInfo>> {
  const resolved = resolveRange("dataValidation.read", sheetName, rangeAddress);
  if (!resolved.ok) return resolved;
  const vResult = requireValidation("dataValidation.read", resolved.data.range);
  if (!vResult.ok) return vResult;
  try {
    const parsed = parseHostValidation(vResult.data, sheetName);
    return ok(toInfo(sheetName, resolved.data.address, parsed));
  } catch (error) {
    return fail("dataValidation.read", "wps-jsa", messageOf(error), EVIDENCE);
  }
}

export async function wpsWriteDataValidation(input: {
  sheetName: string;
  range: string;
  rule: DataValidationRule;
  errorAlert?: unknown;
  prompt?: unknown;
}): Promise<HostResult<DataValidationInfo>> {
  const capability = "dataValidation.write";
  // No in-repo JSA/member evidence for ErrorTitle/ErrorMessage/InputTitle/ShowError etc.
  // Do not invent COM shapes; zero write side effects when new metadata is requested.
  if (input.errorAlert !== undefined || input.prompt !== undefined) {
    return unsupported(
      capability,
      "wps-jsa",
      "errorAlert/prompt require Office.js DataValidation metadata; WPS JSA has no verified ErrorTitle/InputMessage members",
      EVIDENCE,
    );
  }
  const resolved = resolveRange(capability, input.sheetName, input.range);
  if (!resolved.ok) return resolved;
  const vResult = requireValidation(capability, resolved.data.range);
  if (!vResult.ok) return vResult;
  const v = vResult.data;
  if (typeof v.Add !== "function") {
    return unsupported(capability, "wps-jsa", "Validation.Add is unavailable", EVIDENCE);
  }
  const rule = input.rule;
  if (
    rule.type !== "list" &&
    rule.type !== "custom" &&
    !isBetweenOp(rule.operator) &&
    rule.formula2 != null &&
    String(rule.formula2).trim() !== ""
  ) {
    return fail(
      capability,
      "wps-jsa",
      "formula2 is only allowed for between/notBetween",
      EVIDENCE,
    );
  }

  let listSource: string | undefined;
  if (rule.type === "list") {
    const src = buildListSource(rule);
    if (!src.ok) return src;
    listSource = src.data;
  } else if (rule.type === "custom") {
    if (!rule.formula1) {
      return fail(capability, "wps-jsa", "custom requires formula1", EVIDENCE);
    }
  } else if (!rule.operator || !rule.formula1) {
    return fail(
      capability,
      "wps-jsa",
      `${rule.type} requires operator and formula1`,
      EVIDENCE,
    );
  } else if (
    isBetweenOp(rule.operator) &&
    (rule.formula2 == null || String(rule.formula2).trim() === "")
  ) {
    return fail(
      capability,
      "wps-jsa",
      `${rule.type} ${rule.operator} requires formula2`,
      EVIDENCE,
    );
  }

  const snap = trySnapshot(v);
  const hadExisting = snap != null;
  let deleted = false;

  function tryRestore(reason: string): HostResult<DataValidationInfo> {
    if (!(deleted && snap)) {
      return fail(capability, "wps-jsa", reason, EVIDENCE);
    }
    try {
      restoreSnapshot(v, snap);
      return fail(capability, "wps-jsa", `${reason}; original validation restored`, EVIDENCE);
    } catch (restoreError) {
      return fail(
        capability,
        "wps-jsa",
        `${reason}; restore failed: ${messageOf(restoreError)}`,
        EVIDENCE,
      );
    }
  }

  try {
    if (typeof v.Delete === "function") {
      if (hadExisting) {
        try {
          v.Delete();
          deleted = true;
        } catch (error) {
          // Do not call Add after a failed Delete when a prior rule existed.
          return fail(
            capability,
            "wps-jsa",
            `Validation.Delete failed; left original rule in place: ${messageOf(error)}`,
            EVIDENCE,
          );
        }
      } else {
        try {
          v.Delete();
          deleted = true;
        } catch {
          // empty validation Delete may throw; safe to continue Add
        }
      }
    } else if (hadExisting) {
      return unsupported(
        capability,
        "wps-jsa",
        "Validation.Delete is unavailable; cannot safely replace existing rule",
        EVIDENCE,
      );
    }

    if (rule.type === "list") {
      v.Add(XL_VALIDATE_LIST, XL_VALID_ALERT_STOP, undefined, listSource!);
      if (v.InCellDropdown !== undefined) v.InCellDropdown = true;
    } else if (rule.type === "custom") {
      v.Add(XL_VALIDATE_CUSTOM, XL_VALID_ALERT_STOP, undefined, rule.formula1!);
    } else {
      const typeCom = mapDvTypeToCom(rule.type);
      const opCom = mapDvOperatorToCom(rule.operator!);
      if (isBetweenOp(rule.operator)) {
        v.Add(typeCom, XL_VALID_ALERT_STOP, opCom, rule.formula1!, rule.formula2!);
      } else {
        v.Add(typeCom, XL_VALID_ALERT_STOP, opCom, rule.formula1!);
      }
    }
    if (v.IgnoreBlank !== undefined) {
      v.IgnoreBlank = rule.allowBlank !== false;
    }
  } catch (error) {
    return tryRestore(`Validation write failed: ${messageOf(error)}`);
  }

  try {
    const parsed = parseHostValidation(v, input.sheetName);
    if (!parsed.supported || !parsed.rule) {
      return tryRestore(
        `data validation readback not supported after write: ${parsed.hostType}`,
      );
    }
    if (!dvRulesMatch(rule, parsed.rule, parsed.listSourceKind, input.sheetName)) {
      return tryRestore("data validation rule mismatch after write");
    }
    return ok(toInfo(input.sheetName, resolved.data.address, parsed));
  } catch (error) {
    return tryRestore(`Validation readback failed: ${messageOf(error)}`);
  }
}

export async function wpsClearDataValidation(
  sheetName: string,
  rangeAddress: string,
): Promise<HostResult<{ cleared: string }>> {
  const capability = "dataValidation.clear";
  const resolved = resolveRange(capability, sheetName, rangeAddress);
  if (!resolved.ok) return resolved;
  const vResult = requireValidation(capability, resolved.data.range);
  if (!vResult.ok) return vResult;
  const v = vResult.data;
  if (typeof v.Delete !== "function") {
    return unsupported(capability, "wps-jsa", "Validation.Delete is unavailable", EVIDENCE);
  }
  try {
    try {
      v.Delete();
    } catch {
      // already empty
    }
    const parsed = parseHostValidation(v, sheetName);
    if (parsed.hostType !== "None" || parsed.rule != null || parsed.supported) {
      return fail(
        capability,
        "wps-jsa",
        `data validation clear readback not None: hostType=${parsed.hostType}`,
        EVIDENCE,
      );
    }
    return ok({ cleared: resolved.data.address });
  } catch (error) {
    return fail(capability, "wps-jsa", messageOf(error), EVIDENCE);
  }
}
