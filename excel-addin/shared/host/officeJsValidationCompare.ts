/**
 * List-source classification + formula/rule equality for CF/DV readback.
 * Intentionally minimal normalization to avoid false-positive formula matches.
 */
import { isBetweenOp, normalizeHexColor, unmapCfOperator } from "./officeJsValidationMapping";
import type {
  ConditionalFormatRule,
  DataValidationListSourceKind,
  DataValidationRule,
} from "./types";

export type ClassifiedListSource = {
  kind: DataValidationListSourceKind;
  formula1?: string;
  listValues?: string[];
  /** True when source cannot be losslessly written back as supported rule. */
  lossy?: boolean;
  raw?: string;
  limitations?: string[];
};

const SIMPLE_A1_BODY =
  /^(?:'((?:[^']|'')+)'|([A-Za-z0-9_ ]+))!(\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?)$/;
const SIMPLE_A1_BARE = /^(\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?)$/;

function stripOneLeadingEquals(raw: string): string {
  const s = raw.trim();
  return s.startsWith("=") ? s.slice(1).trim() : s;
}

function normalizeSheetKey(sheet: string): string {
  return sheet.replace(/''/g, "'").toLowerCase();
}

/**
 * Writable same-workbook single-area A1 / Sheet!A1(:B2) only.
 * Null for =MyList, =INDIRECT(...), "Yes!,No", multi-area, 3D, structured.
 * Returns { sheet: null for bare, a1 } with a1 uppercased and $ stripped.
 */
export function tryParseSimpleA1Parts(
  raw: string | number | null | undefined,
): { sheet: string | null; a1: string } | null {
  if (raw == null) return null;
  const body = stripOneLeadingEquals(String(raw));
  if (body === "" || body.includes(",") || body.includes("[") || body.includes("]")) {
    return null;
  }
  // Reject function calls / operators / string literals (not pure refs).
  if (/[()"+\-*/^&<>]/.test(body)) return null;

  let sheet: string | null = null;
  let a1: string | undefined;
  const withSheet = SIMPLE_A1_BODY.exec(body);
  if (withSheet) {
    sheet = normalizeSheetKey(withSheet[1] ?? withSheet[2] ?? "");
    a1 = withSheet[3];
    if (!sheet || sheet.includes(":")) return null;
  } else {
    const bare = SIMPLE_A1_BARE.exec(body);
    if (!bare) return null;
    a1 = bare[1];
  }
  if (!a1) return null;
  const bareA1 = a1.replace(/\$/g, "").toUpperCase();
  if (!/^[A-Z]+\d+(:[A-Z]+\d+)?$/.test(bareA1)) return null;
  return { sheet, a1: bareA1 };
}

/** Stable key for simple A1; bare stays bare unless ownerSheetName is provided. */
export function tryNormalizeSimpleA1Ref(
  raw: string | number | null | undefined,
  ownerSheetName?: string,
): string | null {
  const parts = tryParseSimpleA1Parts(raw);
  if (!parts) return null;
  if (parts.sheet != null) return `${parts.sheet}!${parts.a1}`;
  if (ownerSheetName != null && ownerSheetName !== "") {
    return `${normalizeSheetKey(ownerSheetName)}!${parts.a1}`;
  }
  // Bare without owner context — do not invent a sheet qualifier.
  return parts.a1;
}

/**
 * Classify list source string from host readback.
 * - Writable same-workbook A1 → kind=range
 * - "Yes!,No" → inline
 * - =MyList / =INDIRECT(...) → lossy (supported:false upstream)
 */
export function classifyListSource(source: string): ClassifiedListSource {
  const raw = source.trim();
  if (raw === "") {
    return { kind: "inline", listValues: [], lossy: true, raw };
  }

  if (tryParseSimpleA1Parts(raw) != null) {
    return { kind: "range", formula1: stripOneLeadingEquals(raw) };
  }

  // Non-simple leading-= or sheet-like single token → not a writable Range proxy.
  if (raw.startsWith("=") || (!raw.includes(",") && raw.includes("!"))) {
    return {
      kind: "range",
      formula1: raw,
      lossy: true,
      raw,
      limitations: [
        `list source is not a writable same-workbook A1 range (cannot Range-proxy): ${raw}`,
      ],
    };
  }

  const unquoted = raw.replace(/^"(.*)"$/s, "$1");
  const tokens = unquoted.split(",");
  const hasEmptyToken = tokens.some((t) => t.trim() === "");
  const listValues = tokens.map((s) => s.trim()).filter((s) => s.length > 0);
  if (hasEmptyToken) {
    return { kind: "inline", listValues, lossy: true, raw: unquoted };
  }
  return { kind: "inline", listValues, raw: unquoted };
}

/**
 * Minimal formula equality:
 * a) trim + optional single leading = (case-sensitive for non-A1 formulas)
 * b) simple A1 only: strip $ / normalize sheet quoting case;
 *    bare A1 equals Sheet!A1 ONLY when ownerSheetName is provided and sheet is owner.
 * Never lowercases whole formulas or strips $ inside string literals.
 */
export function formulasSemanticallyEqual(
  a: string | number | null | undefined,
  b: string | number | null | undefined,
  ownerSheetName?: string,
): boolean {
  const sa = String(a ?? "").trim();
  const sb = String(b ?? "").trim();
  if (sa === sb) return true;
  const ga = stripOneLeadingEquals(sa);
  const gb = stripOneLeadingEquals(sb);
  if (ga === gb) return true; // "1" vs "=1" only; case-sensitive for EXACT/SUM etc.

  const na = tryNormalizeSimpleA1Ref(sa, ownerSheetName);
  const nb = tryNormalizeSimpleA1Ref(sb, ownerSheetName);
  if (na != null && nb != null) return na === nb;
  return false;
}

export function allowBlankEqual(
  a: boolean | undefined,
  b: boolean | undefined,
): boolean {
  return (a !== false) === (b !== false);
}

export function hostHasExtraFormula2(
  operator: string | undefined,
  formula2: string | number | null | undefined,
): boolean {
  if (isBetweenOp(operator)) return false;
  if (formula2 == null) return false;
  return String(formula2).trim() !== "";
}

export function dvRulesMatch(
  expected: DataValidationRule,
  actual: DataValidationRule,
  actualListKind?: DataValidationListSourceKind | null,
  ownerSheetName?: string,
): boolean {
  if (expected.type !== actual.type) return false;
  if (!allowBlankEqual(expected.allowBlank, actual.allowBlank)) return false;
  if (expected.type === "list") {
    if (expected.listValues && expected.listValues.length > 0) {
      if (actualListKind != null && actualListKind !== "inline") return false;
      const aVals = expected.listValues;
      const bVals = actual.listValues ?? [];
      return aVals.length === bVals.length && aVals.every((v, i) => v === bVals[i]);
    }
    if (expected.formula1) {
      if (actualListKind != null && actualListKind !== "range") return false;
      return formulasSemanticallyEqual(
        expected.formula1,
        actual.formula1,
        ownerSheetName,
      );
    }
    return false;
  }
  if (expected.type === "custom") {
    if (hostHasExtraFormula2(undefined, actual.formula2)) return false;
    return formulasSemanticallyEqual(expected.formula1, actual.formula1, ownerSheetName);
  }
  if (expected.operator !== actual.operator) return false;
  if (!formulasSemanticallyEqual(expected.formula1, actual.formula1, ownerSheetName)) {
    return false;
  }
  if (isBetweenOp(expected.operator)) {
    return formulasSemanticallyEqual(expected.formula2, actual.formula2, ownerSheetName);
  }
  if (hostHasExtraFormula2(expected.operator, actual.formula2)) return false;
  return true;
}

export function cfRuleFieldsMatch(
  expected: ConditionalFormatRule,
  host: {
    operator?: string;
    formula1?: string;
    formula2?: string;
    formula?: string;
    fillColor?: string;
    fontColor?: string;
  },
  ownerSheetName?: string,
): boolean {
  if (expected.kind === "cellValue") {
    if (unmapCfOperator(host.operator) !== expected.operator) return false;
    if (!formulasSemanticallyEqual(expected.formula1, host.formula1, ownerSheetName)) {
      return false;
    }
    if (isBetweenOp(expected.operator)) {
      if (!formulasSemanticallyEqual(expected.formula2, host.formula2, ownerSheetName)) {
        return false;
      }
    } else if (hostHasExtraFormula2(expected.operator, host.formula2)) {
      return false;
    }
  } else if (!formulasSemanticallyEqual(expected.formula, host.formula, ownerSheetName)) {
    return false;
  }
  if (expected.fillColor != null) {
    if (normalizeHexColor(host.fillColor ?? "", "fillColor") !== expected.fillColor) {
      return false;
    }
  }
  if (expected.fontColor != null) {
    if (normalizeHexColor(host.fontColor ?? "", "fontColor") !== expected.fontColor) {
      return false;
    }
  }
  return true;
}
