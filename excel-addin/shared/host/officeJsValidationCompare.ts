/**
 * List-source classification + formula/rule equality for CF/DV readback.
 * Only lossless same-workbook single-area A1 is kind=range.
 */
import { validateBareA1 } from "./officeJsChartSource";
import { isBetweenOp, normalizeHexColor, unmapCfOperator } from "./officeJsValidationMapping";
import type {
  ConditionalFormatRule,
  DataValidationListSourceKind,
  DataValidationRule,
} from "./types";

export type ClassifiedListSource = {
  /** inline | range only when lossless; null = not a writable list source shape. */
  kind: DataValidationListSourceKind | null;
  formula1?: string;
  listValues?: string[];
  lossy?: boolean;
  raw?: string;
  limitations?: string[];
};

/** Excel sheet name forbidden characters (also used in external/3D forms). */
const SHEET_FORBIDDEN = /[\[\]:*?/\\]/;

function stripOneLeadingEquals(raw: string): string {
  const s = raw.trim();
  return s.startsWith("=") ? s.slice(1).trim() : s;
}

function normalizeSheetKey(sheet: string): string {
  return sheet.replace(/''/g, "'").toLowerCase();
}

/**
 * Parse quoted sheet name starting at body[0] === "'".
 * Supports doubled apostrophe; sheet may contain spaces/commas/punctuation.
 */
function parseQuotedSheet(
  body: string,
): { sheet: string; rest: string } | null {
  if (!body.startsWith("'")) return null;
  let i = 1;
  let name = "";
  while (i < body.length) {
    if (body[i] === "'" && body[i + 1] === "'") {
      name += "'";
      i += 2;
      continue;
    }
    if (body[i] === "'") {
      i += 1;
      if (body[i] !== "!") return null;
      if (name === "" || SHEET_FORBIDDEN.test(name) || name.includes(":")) return null;
      return { sheet: name, rest: body.slice(i + 1) };
    }
    name += body[i];
    i += 1;
  }
  return null;
}

/**
 * Unquoted sheet identifier: Unicode letters/numbers/underscore/dot; no spaces.
 * Conservative — never accepts spaces or formula operators.
 */
function isValidUnquotedSheet(name: string): boolean {
  if (name === "" || name.includes(":")) return false;
  if (SHEET_FORBIDDEN.test(name)) return false;
  if (/\s/.test(name)) return false;
  // Letters (any language), numbers, underscore, dot. Must start with letter/underscore.
  return /^[\p{L}_][\p{L}\p{N}_.]*$/u.test(name);
}

function tryValidateA1Body(a1Raw: string): string | null {
  try {
    return validateBareA1(a1Raw, "listSource");
  } catch {
    return null;
  }
}

/**
 * Writable same-workbook single-area A1 / Sheet!A1 / 'Sheet'!A1 only.
 * Null for named ranges, INDIRECT, external, 3D, structured, multi-area, A0, illegal sheets.
 */
export function tryParseSimpleA1Parts(
  raw: string | number | null | undefined,
): { sheet: string | null; a1: string } | null {
  if (raw == null) return null;
  const body = stripOneLeadingEquals(String(raw));
  if (body === "") return null;

  let sheet: string | null = null;
  let a1Part: string;

  if (body.startsWith("'")) {
    const parsed = parseQuotedSheet(body);
    if (!parsed) return null;
    sheet = parsed.sheet;
    a1Part = parsed.rest;
  } else if (body.includes("!")) {
    const bang = body.indexOf("!");
    const sheetRaw = body.slice(0, bang);
    a1Part = body.slice(bang + 1);
    if (!isValidUnquotedSheet(sheetRaw)) return null;
    sheet = sheetRaw;
  } else {
    a1Part = body;
  }

  // A1 body only — no operators/commas left for multi-area after sheet split.
  if (a1Part.includes(",") || a1Part.includes("!") || a1Part.includes("(")) return null;
  const bare = tryValidateA1Body(a1Part);
  if (!bare) return null;
  return { sheet, a1: bare };
}

/** Stable key for simple A1; bare stays bare unless ownerSheetName is provided. */
export function tryNormalizeSimpleA1Ref(
  raw: string | number | null | undefined,
  ownerSheetName?: string,
): string | null {
  const parts = tryParseSimpleA1Parts(raw);
  if (!parts) return null;
  if (parts.sheet != null) return `${normalizeSheetKey(parts.sheet)}!${parts.a1}`;
  if (ownerSheetName != null && ownerSheetName !== "") {
    return `${normalizeSheetKey(ownerSheetName)}!${parts.a1}`;
  }
  return parts.a1;
}

/**
 * True when source is formula-like / illegal range (not a plain inline list).
 * Leading = always unsupported here (simple A1 already handled by caller).
 */
function isUnsupportedFormulaLikeListSource(raw: string): boolean {
  if (raw.startsWith("=")) return true;
  if (raw.includes("[") || raw.includes("]")) return true;
  // 3D unquoted Sheet1:Sheet3!A1
  if (!raw.startsWith("'") && /^.+:.+!/.test(raw)) return true;

  if (!raw.includes("!")) return false;

  // Multi-token: multi-area of A1/sheet!A1 forms → unsupported (not plain CSV values).
  if (raw.includes(",")) {
    const tokens = raw.split(",").map((t) => t.trim()).filter((t) => t.length > 0);
    for (const token of tokens) {
      if (tryParseSimpleA1Parts(token) != null) return true;
      if (token.includes("!")) {
        const after = token.slice(token.indexOf("!") + 1);
        if (tryValidateA1Body(after) != null) return true;
      }
    }
    return false; // e.g. Yes!,No
  }

  // Single token with ! that is not simple A1: failed sheet!A1 form → unsupported
  // when the part after ! is empty or A1-like (Sheet 1!A1, Sheet! , etc.).
  const after = raw.slice(raw.indexOf("!") + 1);
  if (after.trim() === "" || tryValidateA1Body(after) != null) return true;
  return false;
}

function parseInlineListSource(raw: string): ClassifiedListSource {
  const unquoted = raw.replace(/^"(.*)"$/s, "$1");
  const tokens = unquoted.split(",");
  const hasEmptyToken = tokens.some((t) => t.trim() === "");
  const listValues = tokens.map((s) => s.trim()).filter((s) => s.length > 0);
  if (listValues.length === 0) {
    return { kind: "inline", listValues: [], lossy: true, raw: unquoted };
  }
  if (hasEmptyToken) {
    return { kind: "inline", listValues, lossy: true, raw: unquoted };
  }
  return { kind: "inline", listValues, raw: unquoted };
}

/**
 * Classify list source string from host readback.
 * 1) lossless simple A1 → kind=range
 * 2) formula-like / illegal range → kind=null (never kind=range)
 * 3) remaining non-empty strings → inline (single or multi token; with or without commas)
 */
export function classifyListSource(source: string): ClassifiedListSource {
  const raw = source.trim();
  if (raw === "") {
    return { kind: "inline", listValues: [], lossy: true, raw };
  }

  if (tryParseSimpleA1Parts(raw) != null) {
    return { kind: "range", formula1: stripOneLeadingEquals(raw) };
  }

  if (isUnsupportedFormulaLikeListSource(raw)) {
    return {
      kind: null,
      formula1: raw,
      lossy: true,
      raw,
      limitations: [
        `list source is not a writable same-workbook A1 range (cannot Range-proxy): ${raw}`,
      ],
    };
  }

  return parseInlineListSource(raw);
}

/**
 * Minimal formula equality:
 * a) trim + optional single leading = (case-sensitive for non-A1 formulas)
 * b) simple A1 only with owner-scoped bare ≡ owner!A1
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
  if (ga === gb) return true;

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
