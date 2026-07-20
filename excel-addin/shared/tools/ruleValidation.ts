import type {
  CellValueOperator,
  ConditionalFormatRule,
  DataValidationAlertStyle,
  DataValidationErrorAlert,
  DataValidationOperator,
  DataValidationPrompt,
  DataValidationRule,
  DataValidationType,
} from "../host/types";
import {
  DV_ALERT_STYLES,
  MAX_DV_ALERT_MESSAGE_CHARS,
  MAX_DV_ALERT_TITLE_CHARS,
} from "../host/officeJsValidationAlerts";
import {
  isBetweenOp,
  MAX_INLINE_LIST_SOURCE_CHARS,
  MAX_LIST_VALUES,
  optionalHexColor,
} from "../host/officeJsValidationMapping";

const CF_OPS: readonly CellValueOperator[] = [
  "greaterThan",
  "greaterThanOrEqualTo",
  "lessThan",
  "lessThanOrEqualTo",
  "equalTo",
  "notEqualTo",
  "between",
  "notBetween",
];

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

const DV_TYPES: readonly DataValidationType[] = [
  "list",
  "wholeNumber",
  "decimal",
  "date",
  "time",
  "textLength",
  "custom",
];

const CF_ALLOWED = new Set([
  "kind",
  "operator",
  "formula1",
  "formula2",
  "formula",
  "fillColor",
  "fontColor",
]);

const DV_ALLOWED = new Set([
  "type",
  "operator",
  "formula1",
  "formula2",
  "listValues",
  "allowBlank",
]);

function requireNonEmptyString(value: unknown, field: string): string {
  if (typeof value !== "string") throw new Error(`${field} must be a string`);
  if (value.trim() === "") throw new Error(`${field} must be non-empty`);
  return value;
}

function parseAllowBlank(raw: unknown): boolean | undefined {
  if (raw === undefined) return undefined;
  if (typeof raw !== "boolean") throw new Error("allowBlank must be a boolean");
  return raw;
}

function rejectFormula2IfNotBetween(operator: string, formula2: unknown): void {
  if (formula2 !== undefined && formula2 !== null) {
    throw new Error(`${operator} must not include formula2`);
  }
}

export function requireCfRule(args: Record<string, unknown>): ConditionalFormatRule {
  const rule = args.rule;
  if (!rule || typeof rule !== "object" || Array.isArray(rule)) {
    throw new Error("rule must be an object");
  }
  const r = rule as Record<string, unknown>;
  if (r.kind !== "cellValue" && r.kind !== "custom") {
    throw new Error('rule.kind must be "cellValue" or "custom"');
  }
  for (const key of Object.keys(r)) {
    if (!CF_ALLOWED.has(key)) throw new Error(`unknown rule field: ${key}`);
  }
  const fillColor = optionalHexColor(r.fillColor, "fillColor");
  const fontColor = optionalHexColor(r.fontColor, "fontColor");

  if (r.kind === "cellValue") {
    if (r.formula != null) throw new Error("cellValue must not include formula");
    if (typeof r.operator !== "string" || !CF_OPS.includes(r.operator as CellValueOperator)) {
      throw new Error(
        "cellValue requires operator greaterThan|greaterThanOrEqualTo|lessThan|lessThanOrEqualTo|equalTo|notEqualTo|between|notBetween",
      );
    }
    const formula1 = requireNonEmptyString(r.formula1, "formula1");
    let formula2: string | undefined;
    if (isBetweenOp(r.operator)) {
      formula2 = requireNonEmptyString(r.formula2, "formula2");
    } else {
      rejectFormula2IfNotBetween(r.operator, r.formula2);
    }
    return {
      kind: "cellValue",
      operator: r.operator as CellValueOperator,
      formula1,
      formula2,
      fillColor,
      fontColor,
    };
  }

  if (r.operator != null) throw new Error("custom must not include operator");
  if (r.formula1 != null) throw new Error("custom must not include formula1");
  if (r.formula2 != null) throw new Error("custom must not include formula2");
  const formula = requireNonEmptyString(r.formula, "formula");
  return { kind: "custom", formula, fillColor, fontColor };
}

export function requireDvRule(args: Record<string, unknown>): DataValidationRule {
  const rule = args.rule;
  if (!rule || typeof rule !== "object" || Array.isArray(rule)) {
    throw new Error("rule must be an object");
  }
  const r = rule as Record<string, unknown>;
  if (typeof r.type !== "string" || !DV_TYPES.includes(r.type as DataValidationType)) {
    throw new Error(
      "rule.type must be list|wholeNumber|decimal|date|time|textLength|custom",
    );
  }
  for (const key of Object.keys(r)) {
    if (!DV_ALLOWED.has(key)) throw new Error(`unknown rule field: ${key}`);
  }
  const allowBlank = parseAllowBlank(r.allowBlank);

  if (r.type === "list") {
    const hasList = Array.isArray(r.listValues);
    const hasFormula = r.formula1 != null && r.formula1 !== "";
    if (hasList && hasFormula) {
      throw new Error("list must not combine listValues and formula1 (mutually exclusive)");
    }
    if (hasList) {
      const listValues = r.listValues as unknown[];
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
      const serialized = out.join(",");
      if (serialized.length > MAX_INLINE_LIST_SOURCE_CHARS) {
        throw new Error(
          `inline list source exceeds Excel ${MAX_INLINE_LIST_SOURCE_CHARS} character limit; use a range source`,
        );
      }
      if (r.operator != null) throw new Error("list must not include operator");
      if (r.formula2 != null) throw new Error("list must not include formula2");
      return { type: "list", listValues: out, allowBlank };
    }
    if (hasFormula) {
      if (typeof r.formula1 !== "string") throw new Error("formula1 must be a string");
      requireNonEmptyString(r.formula1, "formula1");
      if (r.operator != null) throw new Error("list must not include operator");
      if (r.formula2 != null) throw new Error("list must not include formula2");
      return { type: "list", formula1: r.formula1, allowBlank };
    }
    throw new Error("list requires listValues or formula1 range source");
  }

  if (r.type === "custom") {
    if (r.operator != null) throw new Error("custom must not include operator");
    if (r.formula2 != null) throw new Error("custom must not include formula2");
    if (r.listValues != null) throw new Error("custom must not include listValues");
    const formula1 = requireNonEmptyString(r.formula1, "formula1");
    return { type: "custom", formula1, allowBlank };
  }

  if (typeof r.operator !== "string" || !DV_OPS.includes(r.operator as DataValidationOperator)) {
    throw new Error(
      `${r.type} requires operator between|notBetween|equalTo|notEqualTo|greaterThan|greaterThanOrEqualTo|lessThan|lessThanOrEqualTo`,
    );
  }
  if (r.listValues != null) throw new Error(`${r.type} must not include listValues`);
  const formula1 = requireNonEmptyString(r.formula1, "formula1");
  let formula2: string | undefined;
  if (isBetweenOp(r.operator)) {
    formula2 = requireNonEmptyString(r.formula2, "formula2");
  } else {
    rejectFormula2IfNotBetween(r.operator, r.formula2);
  }
  return {
    type: r.type as DataValidationType,
    operator: r.operator as DataValidationOperator,
    formula1,
    formula2,
    allowBlank,
  };
}

const ERROR_ALERT_KEYS = new Set(["showAlert", "style", "title", "message"]);
const PROMPT_KEYS = new Set(["showPrompt", "title", "message"]);

function requireAlertString(raw: unknown, field: string, max: number): string {
  if (typeof raw !== "string") throw new Error(`${field} must be a string`);
  if (raw.length > max) {
    throw new Error(`${field} maxLength is ${max}`);
  }
  return raw;
}

export function requireDvErrorAlert(raw: unknown): DataValidationErrorAlert {
  if (raw === null) throw new Error("errorAlert must not be null");
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) {
    throw new Error("errorAlert must be an object");
  }
  const o = raw as Record<string, unknown>;
  for (const key of Object.keys(o)) {
    if (!ERROR_ALERT_KEYS.has(key)) throw new Error(`unknown errorAlert field: ${key}`);
  }
  const out: DataValidationErrorAlert = {};
  if (o.showAlert !== undefined) {
    if (typeof o.showAlert !== "boolean") throw new Error("errorAlert.showAlert must be a boolean");
    out.showAlert = o.showAlert;
  }
  if (o.style !== undefined) {
    if (typeof o.style !== "string" || !(DV_ALERT_STYLES as readonly string[]).includes(o.style)) {
      throw new Error("errorAlert.style must be stop|warning|information");
    }
    out.style = o.style as DataValidationAlertStyle;
  }
  if (o.title !== undefined) {
    out.title = requireAlertString(o.title, "errorAlert.title", MAX_DV_ALERT_TITLE_CHARS);
  }
  if (o.message !== undefined) {
    out.message = requireAlertString(o.message, "errorAlert.message", MAX_DV_ALERT_MESSAGE_CHARS);
  }
  return out;
}

export function requireDvPrompt(raw: unknown): DataValidationPrompt {
  if (raw === null) throw new Error("prompt must not be null");
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) {
    throw new Error("prompt must be an object");
  }
  const o = raw as Record<string, unknown>;
  for (const key of Object.keys(o)) {
    if (!PROMPT_KEYS.has(key)) throw new Error(`unknown prompt field: ${key}`);
  }
  const out: DataValidationPrompt = {};
  if (o.showPrompt !== undefined) {
    if (typeof o.showPrompt !== "boolean") throw new Error("prompt.showPrompt must be a boolean");
    out.showPrompt = o.showPrompt;
  }
  if (o.title !== undefined) {
    out.title = requireAlertString(o.title, "prompt.title", MAX_DV_ALERT_TITLE_CHARS);
  }
  if (o.message !== undefined) {
    out.message = requireAlertString(o.message, "prompt.message", MAX_DV_ALERT_MESSAGE_CHARS);
  }
  return out;
}

export function optionalDvErrorAlert(
  args: Record<string, unknown>,
): DataValidationErrorAlert | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "errorAlert")) return undefined;
  return requireDvErrorAlert(args.errorAlert);
}

export function optionalDvPrompt(
  args: Record<string, unknown>,
): DataValidationPrompt | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "prompt")) return undefined;
  return requireDvPrompt(args.prompt);
}
