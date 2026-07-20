/**
 * CF/DV official host type & operator maps (ExcelApi 1.6 CF / 1.8 DV).
 * CF GTE/LTE host tokens omit "To"; DV tokens include "To" — never share maps.
 * CF notEqualTo host token is NotEqualTo (official ConditionalCellValueOperator).
 */
import type {
  CellValueOperator,
  ConditionalFormatListKind,
  DataValidationOperator,
  DataValidationType,
} from "./types";

/** Official ConditionalFormatType values we recognize on list. */
export const CF_HOST_TYPES = [
  "CellValue",
  "Custom",
  "DataBar",
  "ColorScale",
  "IconSet",
  "TopBottom",
  "PresetCriteria",
  "ContainsText",
] as const;

/** Public CF operators → ConditionalCellValueRule.operator host tokens. */
export const CF_OP_TO_HOST: Record<CellValueOperator, string> = {
  greaterThan: "GreaterThan",
  greaterThanOrEqualTo: "GreaterThanOrEqual",
  lessThan: "LessThan",
  lessThanOrEqualTo: "LessThanOrEqual",
  equalTo: "EqualTo",
  notEqualTo: "NotEqualTo",
  between: "Between",
  notBetween: "NotBetween",
};

const CF_HOST_OP_TO_PUBLIC: Record<string, CellValueOperator> = {
  greaterthan: "greaterThan",
  greaterthanorequal: "greaterThanOrEqualTo",
  lessthan: "lessThan",
  lessthanorequal: "lessThanOrEqualTo",
  equalto: "equalTo",
  notequalto: "notEqualTo",
  between: "between",
  notbetween: "notBetween",
};

/** Public DV operators → DataValidationOperator host tokens (…OrEqualTo). */
export const DV_OP_TO_HOST: Record<DataValidationOperator, string> = {
  greaterThan: "GreaterThan",
  greaterThanOrEqualTo: "GreaterThanOrEqualTo",
  lessThan: "LessThan",
  lessThanOrEqualTo: "LessThanOrEqualTo",
  equalTo: "EqualTo",
  notEqualTo: "NotEqualTo",
  between: "Between",
  notBetween: "NotBetween",
};

const DV_HOST_OP_TO_PUBLIC: Record<string, DataValidationOperator> = {
  greaterthan: "greaterThan",
  greaterthanorequalto: "greaterThanOrEqualTo",
  lessthan: "lessThan",
  lessthanorequalto: "lessThanOrEqualTo",
  equalto: "equalTo",
  notequalto: "notEqualTo",
  between: "between",
  notbetween: "notBetween",
};

export const COMPARE_DV_TYPES: readonly DataValidationType[] = [
  "wholeNumber",
  "decimal",
  "date",
  "time",
  "textLength",
];

export const MAX_LIST_VALUES = 1000;
/** Excel native inline list source max length (characters). */
export const MAX_INLINE_LIST_SOURCE_CHARS = 255;

export function normalizeToken(raw: string): string {
  return String(raw ?? "").replace(/\s+/g, "").toLowerCase();
}

export function mapCfOperatorToHost(op: CellValueOperator): string {
  return CF_OP_TO_HOST[op];
}

/** Exact token map — never use includes() (would map >= as >). */
export function unmapCfOperator(host: string | null | undefined): CellValueOperator | undefined {
  return CF_HOST_OP_TO_PUBLIC[normalizeToken(String(host ?? ""))];
}

export function mapDvOperatorToHost(op: DataValidationOperator): string {
  return DV_OP_TO_HOST[op];
}

export function unmapDvOperator(host: string | null | undefined): DataValidationOperator | undefined {
  return DV_HOST_OP_TO_PUBLIC[normalizeToken(String(host ?? ""))];
}

export function classifyCfHostType(raw: string | null | undefined): {
  kind: ConditionalFormatListKind;
  hostType: string;
  supported: boolean;
  limitations?: string[];
} {
  const token = normalizeToken(String(raw ?? ""));
  if (token === "cellvalue") {
    return { kind: "cellValue", hostType: "CellValue", supported: true };
  }
  if (token === "custom") {
    return { kind: "custom", hostType: "Custom", supported: true };
  }
  const known: Record<string, string> = {
    databars: "DataBar",
    databar: "DataBar",
    colorscale: "ColorScale",
    iconset: "IconSet",
    topbottom: "TopBottom",
    presetcriteria: "PresetCriteria",
    containstext: "ContainsText",
  };
  const hostType = known[token] ?? (String(raw ?? "Unknown") || "Unknown");
  return {
    kind: "unsupported",
    hostType,
    supported: false,
    limitations: [
      `ConditionalFormatType ${hostType} is recognized but not add-capable in this add-in (list-only honesty)`,
    ],
  };
}

export function classifyDvHostType(raw: string | null | undefined): {
  type: DataValidationType | null;
  hostType: string;
  writable: boolean;
  mixedState: boolean;
  limitations?: string[];
} {
  const token = normalizeToken(String(raw ?? ""));
  if (!token || token === "none" || token === "null" || token === "undefined") {
    return { type: null, hostType: "None", writable: false, mixedState: false };
  }
  if (token === "inconsistent" || token === "mixedcriteria") {
    const hostType = token === "inconsistent" ? "Inconsistent" : "MixedCriteria";
    return {
      type: null,
      hostType,
      writable: false,
      mixedState: true,
      limitations: [
        `DataValidation type is ${hostType}: region has mixed/partial rules — not a single writable rule`,
      ],
    };
  }
  const map: Record<string, DataValidationType> = {
    list: "list",
    wholenumber: "wholeNumber",
    whole: "wholeNumber",
    decimal: "decimal",
    date: "date",
    time: "time",
    textlength: "textLength",
    custom: "custom",
  };
  const type = map[token];
  if (type) {
    return { type, hostType: hostTypeLabel(type), writable: true, mixedState: false };
  }
  return {
    type: null,
    hostType: String(raw ?? "Unknown"),
    writable: false,
    mixedState: false,
    limitations: [
      `Unknown DataValidationType ${String(raw)} — not coerced to wholeNumber/list`,
    ],
  };
}

function hostTypeLabel(type: DataValidationType): string {
  switch (type) {
    case "wholeNumber":
      return "WholeNumber";
    case "textLength":
      return "TextLength";
    case "list":
      return "List";
    case "decimal":
      return "Decimal";
    case "date":
      return "Date";
    case "time":
      return "Time";
    case "custom":
      return "Custom";
  }
}

export function isBetweenOp(op: string | undefined): boolean {
  return op === "between" || op === "notBetween";
}

/** #RRGGBB only; optional bare RRGGBB normalized. */
export function normalizeHexColor(raw: unknown, field: string): string {
  if (typeof raw !== "string") throw new Error(`${field} must be a string (#RRGGBB)`);
  const hex = raw.startsWith("#") ? raw.slice(1) : raw;
  if (!/^[0-9A-Fa-f]{6}$/.test(hex)) {
    throw new Error(`${field} must be #RRGGBB`);
  }
  return `#${hex.toUpperCase()}`;
}

/** Missing field → undefined; empty string is invalid (must error). */
export function optionalHexColor(raw: unknown, field: string): string | undefined {
  if (raw === undefined || raw === null) return undefined;
  if (raw === "") throw new Error(`${field} must be #RRGGBB (empty string not allowed)`);
  return normalizeHexColor(raw, field);
}
