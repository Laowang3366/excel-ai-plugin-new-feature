/**
 * Pure Office.js DataValidation rule planner.
 * Builds the final host rule object before any dv.rule / ignoreBlanks / alert assignment.
 */
import type { ExcelDataValidationRule } from "./officeJsExcelTypes";
import type { ExcelRequestContext } from "./officeJsRuntime";
import { resolveListSourceRange } from "./officeJsValidationListSource";
import {
  COMPARE_DV_TYPES,
  isBetweenOp,
  mapDvOperatorToHost,
  MAX_INLINE_LIST_SOURCE_CHARS,
  MAX_LIST_VALUES,
} from "./officeJsValidationMapping";
import type {
  DataValidationOperator,
  DataValidationRule,
  DataValidationType,
} from "./types";

const DV_OPS: readonly DataValidationOperator[] = [
  "between",
  "notBetween",
  "equalTo",
  "notEqualTo",
  "greaterThan",
  "greaterThanOrEqualTo",
  "lessThan",
  "lessThanOrEqualTo",
];

const WRITABLE_TYPES: readonly DataValidationType[] = [
  "list",
  "wholeNumber",
  "decimal",
  "date",
  "time",
  "textLength",
  "custom",
];

function requireNonEmptyString(value: unknown, field: string): string {
  if (typeof value !== "string") throw new Error(`${field} must be a string`);
  if (value.trim() === "") throw new Error(`${field} must be non-empty`);
  return value;
}

/**
 * Validate public rule and construct the final host ExcelDataValidationRule.
 * Must complete without mutating DataValidation; throws on any illegal combination.
 */
export function planDvRuleWrite(
  context: ExcelRequestContext,
  ownerSheetName: string,
  rule: DataValidationRule,
): ExcelDataValidationRule {
  if (!rule || typeof rule !== "object") {
    throw new Error("rule must be an object");
  }
  if (typeof rule.type !== "string" || !(WRITABLE_TYPES as readonly string[]).includes(rule.type)) {
    throw new Error(
      "rule.type must be list|wholeNumber|decimal|date|time|textLength|custom",
    );
  }

  if (rule.type === "list") {
    if (rule.operator != null) throw new Error("list must not include operator");
    if (rule.formula2 != null) throw new Error("list must not include formula2");
    const hasList = Array.isArray(rule.listValues);
    const hasFormula =
      rule.formula1 != null && String(rule.formula1).trim() !== "";
    if (hasList && hasFormula) {
      throw new Error("list must not combine listValues and formula1 (mutually exclusive)");
    }
    if (!hasList && !hasFormula) {
      throw new Error("list requires listValues or formula1 range source");
    }
    if (hasList) {
      const listValues = rule.listValues as unknown[];
      if (listValues.length === 0) throw new Error("listValues must be non-empty");
      if (listValues.length > MAX_LIST_VALUES) {
        throw new Error(`listValues maxItems is ${MAX_LIST_VALUES}`);
      }
      const out: string[] = [];
      for (const item of listValues) {
        if (typeof item !== "string" || item.trim() === "") {
          throw new Error("listValues items must be non-empty strings");
        }
        if (item.includes(",")) {
          throw new Error("listValues items must not contain commas; use a range source");
        }
        out.push(item);
      }
      const source = out.join(",");
      if (source.length > MAX_INLINE_LIST_SOURCE_CHARS) {
        throw new Error(
          `inline list source exceeds Excel ${MAX_INLINE_LIST_SOURCE_CHARS} character limit; use a range source`,
        );
      }
      return { list: { inCellDropDown: true, source } };
    }
    const formula1 = requireNonEmptyString(rule.formula1, "formula1");
    const rangeSource = resolveListSourceRange(context, ownerSheetName, formula1);
    return { list: { inCellDropDown: true, source: rangeSource } };
  }

  if (rule.type === "custom") {
    if (rule.operator != null) throw new Error("custom must not include operator");
    if (rule.formula2 != null) throw new Error("custom must not include formula2");
    if (rule.listValues != null) throw new Error("custom must not include listValues");
    const formula1 = requireNonEmptyString(rule.formula1, "formula1");
    return { custom: { formula: formula1 } };
  }

  // Compare types
  if (!(COMPARE_DV_TYPES as readonly string[]).includes(rule.type)) {
    throw new Error(`unsupported data validation type: ${rule.type}`);
  }
  if (rule.listValues != null) {
    throw new Error(`${rule.type} must not include listValues`);
  }
  if (
    typeof rule.operator !== "string" ||
    !(DV_OPS as readonly string[]).includes(rule.operator)
  ) {
    throw new Error(
      `${rule.type} requires operator between|notBetween|equalTo|notEqualTo|greaterThan|greaterThanOrEqualTo|lessThan|lessThanOrEqualTo`,
    );
  }
  const formula1 = requireNonEmptyString(rule.formula1, "formula1");
  const bag: {
    formula1: string;
    formula2?: string;
    operator: string;
  } = {
    formula1,
    operator: mapDvOperatorToHost(rule.operator as DataValidationOperator),
  };
  if (isBetweenOp(rule.operator)) {
    bag.formula2 = requireNonEmptyString(rule.formula2, "formula2");
  } else if (rule.formula2 !== undefined && rule.formula2 !== null) {
    // including empty string — explicit formula2 is illegal for non-between
    throw new Error(`${rule.operator} must not include formula2`);
  }
  return { [rule.type]: bag } as ExcelDataValidationRule;
}
