import type { ConditionalFormatRule, DataValidationRule } from "../host/types";

const CF_OPS = ["greaterThan", "lessThan", "equalTo", "between", "notBetween"] as const;
const DV_OPS = ["between", "notBetween", "equalTo", "greaterThan", "lessThan"] as const;

export function requireCfRule(args: Record<string, unknown>): ConditionalFormatRule {
  const rule = args.rule;
  if (!rule || typeof rule !== "object" || Array.isArray(rule)) {
    throw new Error("rule must be an object");
  }
  const r = rule as Record<string, unknown>;
  if (r.kind !== "cellValue" && r.kind !== "custom") {
    throw new Error('rule.kind must be "cellValue" or "custom"');
  }
  const allowed = new Set([
    "kind",
    "operator",
    "formula1",
    "formula2",
    "formula",
    "fillColor",
    "fontColor",
  ]);
  for (const key of Object.keys(r)) {
    if (!allowed.has(key)) throw new Error(`unknown rule field: ${key}`);
  }
  if (r.kind === "cellValue") {
    if (typeof r.operator !== "string" || !CF_OPS.includes(r.operator as (typeof CF_OPS)[number])) {
      throw new Error("cellValue requires operator greaterThan|lessThan|equalTo|between|notBetween");
    }
    if (typeof r.formula1 !== "string" || r.formula1.trim() === "") {
      throw new Error("cellValue requires formula1");
    }
    if (
      (r.operator === "between" || r.operator === "notBetween") &&
      (typeof r.formula2 !== "string" || r.formula2.trim() === "")
    ) {
      throw new Error("between/notBetween requires formula2");
    }
  } else if (typeof r.formula !== "string" || r.formula.trim() === "") {
    throw new Error("custom requires formula");
  }
  return r as unknown as ConditionalFormatRule;
}

export function requireDvRule(args: Record<string, unknown>): DataValidationRule {
  const rule = args.rule;
  if (!rule || typeof rule !== "object" || Array.isArray(rule)) {
    throw new Error("rule must be an object");
  }
  const r = rule as Record<string, unknown>;
  if (r.type !== "list" && r.type !== "wholeNumber") {
    throw new Error('rule.type must be "list" or "wholeNumber"');
  }
  // Implemented fields only — showError/errorMessage are not wired to Office.js errorAlert
  const allowed = new Set([
    "type",
    "operator",
    "formula1",
    "formula2",
    "listValues",
    "allowBlank",
  ]);
  for (const key of Object.keys(r)) {
    if (!allowed.has(key)) throw new Error(`unknown rule field: ${key}`);
  }
  if (r.type === "list") {
    if (Array.isArray(r.listValues)) {
      if (r.listValues.length === 0) throw new Error("listValues must be non-empty");
      for (const item of r.listValues) {
        if (typeof item !== "string" || item.trim() === "") {
          throw new Error("listValues items must be non-empty strings");
        }
      }
    } else if (typeof r.formula1 !== "string" || r.formula1.trim() === "") {
      throw new Error("list requires listValues or formula1");
    }
  } else {
    if (typeof r.operator !== "string" || !DV_OPS.includes(r.operator as (typeof DV_OPS)[number])) {
      throw new Error(
        "wholeNumber requires operator between|notBetween|equalTo|greaterThan|lessThan",
      );
    }
    if (typeof r.formula1 !== "string" || r.formula1.trim() === "") {
      throw new Error("wholeNumber requires formula1");
    }
    if (
      (r.operator === "between" || r.operator === "notBetween") &&
      (typeof r.formula2 !== "string" || r.formula2.trim() === "")
    ) {
      throw new Error("between/notBetween requires formula2");
    }
  }
  return r as unknown as DataValidationRule;
}
