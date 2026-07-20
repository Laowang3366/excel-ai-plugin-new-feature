/**
 * Portable subset of desktop ExcelFormulaClassification.
 * Used for write-intent hints; host still decides Formula vs FormulaArray support.
 */

import type { FormulaKind } from "./types";

const FUNCTION_HEAD_RE = /^\s*([A-Za-z_][A-Za-z0-9_.]*)\s*\(/;

const RANGE_REF_RE =
  /(?<![A-Za-z0-9_.])(?:\$?[A-Za-z]{1,3}\$?\d+\s*:\s*\$?[A-Za-z]{1,3}\$?\d+|\$?[A-Za-z]{1,3}\s*:\s*\$?[A-Za-z]{1,3}|\$?\d+\s*:\s*\$?\d+)/i;

const LEGACY_SPILL = new Set([
  "TRANSPOSE",
  "FREQUENCY",
  "MMULT",
  "LINEST",
  "LOGEST",
  "TREND",
  "GROWTH",
]);

const SCALAR_RANGE = new Set([
  "SUM",
  "SUMIF",
  "SUMIFS",
  "COUNT",
  "COUNTA",
  "COUNTBLANK",
  "COUNTIF",
  "COUNTIFS",
  "AVERAGE",
  "AVERAGEIF",
  "AVERAGEIFS",
  "MIN",
  "MAX",
  "MEDIAN",
  "PRODUCT",
  "SUBTOTAL",
  "AGGREGATE",
  "LOOKUP",
  "VLOOKUP",
  "HLOOKUP",
  "MATCH",
  "INDEX",
  "AND",
  "OR",
]);

/** Modern dynamic-array / lambda family (desktop FunctionPrefixes keys). */
const MODERN_FUNCTIONS = new Set([
  "LET",
  "LAMBDA",
  "XLOOKUP",
  "XMATCH",
  "TAKE",
  "DROP",
  "CHOOSEROWS",
  "CHOOSECOLS",
  "MAP",
  "REDUCE",
  "SCAN",
  "HSTACK",
  "VSTACK",
  "MAKEARRAY",
  "BYROW",
  "BYCOL",
  "TEXTSPLIT",
  "TEXTBEFORE",
  "TEXTAFTER",
  "FILTER",
  "SORT",
  "SORTBY",
  "UNIQUE",
  "SEQUENCE",
  "RANDARRAY",
  "TOCOL",
  "TOROW",
  "WRAPROWS",
  "WRAPCOLS",
  "GROUPBY",
  "PIVOTBY",
  "EXPAND",
]);

export function isFormula(text: string | null | undefined): boolean {
  return typeof text === "string" && text.length > 1 && text.startsWith("=");
}

function bareFunctionName(qualified: string): string {
  const dot = qualified.lastIndexOf(".");
  return (dot >= 0 ? qualified.slice(dot + 1) : qualified).toUpperCase();
}

export function leadingFunction(formula: string | null | undefined): string | null {
  if (!isFormula(formula)) return null;
  const match = FUNCTION_HEAD_RE.exec(formula!.slice(1));
  if (!match) return null;
  return bareFunctionName(match[1] ?? "");
}

/** Blank string + sheet quotes so classification ignores quoted text. */
function removeQuotedContent(text: string): string {
  const output = text.split("");
  for (let index = 0; index < output.length; ) {
    const ch = output[index];
    if (ch !== '"' && ch !== "'") {
      index += 1;
      continue;
    }
    const quote = ch;
    output[index] = " ";
    index += 1;
    while (index < output.length) {
      const current = output[index]!;
      output[index] = " ";
      index += 1;
      if (current !== quote) continue;
      if (index < output.length && output[index] === quote) {
        output[index] = " ";
        index += 1;
        continue;
      }
      break;
    }
  }
  return output.join("");
}

function skipStringLiteral(body: string, index: { i: number }): void {
  index.i += 1;
  while (index.i < body.length) {
    if (body[index.i++] !== '"') continue;
    if (index.i < body.length && body[index.i] === '"') {
      index.i += 1;
      continue;
    }
    return;
  }
}

function skipQuotedSheet(body: string, index: { i: number }): void {
  index.i += 1;
  while (index.i < body.length) {
    if (body[index.i++] !== "'") continue;
    if (index.i < body.length && body[index.i] === "'") {
      index.i += 1;
      continue;
    }
    return;
  }
}

function isIdentStart(ch: string): boolean {
  return /[A-Za-z_]/.test(ch);
}

function isIdentPart(ch: string): boolean {
  return /[A-Za-z0-9_.]/.test(ch);
}

function containsModernFunctionCall(formula: string): boolean {
  const body = formula.slice(1);
  const index = { i: 0 };
  while (index.i < body.length) {
    const ch = body[index.i]!;
    if (ch === '"') {
      skipStringLiteral(body, index);
      continue;
    }
    if (ch === "'") {
      skipQuotedSheet(body, index);
      continue;
    }
    if (!isIdentStart(ch)) {
      index.i += 1;
      continue;
    }
    const start = index.i;
    index.i += 1;
    while (index.i < body.length && isIdentPart(body[index.i]!)) index.i += 1;
    const qualified = body.slice(start, index.i);
    let next = index.i;
    while (next < body.length && /\s/.test(body[next]!)) next += 1;
    const fn = bareFunctionName(qualified);
    if (next < body.length && body[next] === "(" && MODERN_FUNCTIONS.has(fn)) {
      return true;
    }
  }
  return false;
}

function containsLegacySpillFunction(formula: string): boolean {
  const lead = leadingFunction(formula);
  return lead != null && LEGACY_SPILL.has(lead);
}

function containsArrayExpression(formula: string): boolean {
  const searchable = removeQuotedContent(formula.slice(1));
  if (!RANGE_REF_RE.test(searchable)) return false;
  const lead = leadingFunction(formula);
  return lead == null || !SCALAR_RANGE.has(lead);
}

export function isDynamicArray(
  formula: string | null | undefined,
  forceLegacyArray = false,
): boolean {
  if (forceLegacyArray || !isFormula(formula)) return false;
  return (
    containsModernFunctionCall(formula!) ||
    containsLegacySpillFunction(formula!) ||
    containsArrayExpression(formula!)
  );
}

/**
 * Classify formula write intent.
 * @throws if text is not a formula (desktop throws ArgumentException).
 */
export function classifyFormula(
  formula: string,
  legacyCse = false,
): FormulaKind {
  if (!isFormula(formula)) {
    throw new Error("公式必须以 '=' 开头");
  }
  if (legacyCse) return "legacyArray";
  return isDynamicArray(formula) ? "dynamic" : "plain";
}

/** Soft classify: non-formulas → null. */
export function tryClassifyFormula(
  formula: string | null | undefined,
  legacyCse = false,
): FormulaKind | null {
  if (!isFormula(formula)) return null;
  if (legacyCse) return "legacyArray";
  return isDynamicArray(formula!) ? "dynamic" : "plain";
}
