/**
 * Excel/WPS COM constants for FormatConditions + Validation.
 * Desktop ExcelActionService uses dynamic FormatConditions.Add / Validation.Add.
 */
import type {
  CellValueOperator,
  ConditionalFormatListKind,
  DataValidationOperator,
  DataValidationType,
} from "./types";

/** FormatConditionType */
export const XL_CELL_VALUE = 1;
export const XL_EXPRESSION = 2;

/** xlFormatConditionOperator / xlDVOperator */
export const XL_BETWEEN = 1;
export const XL_NOT_BETWEEN = 2;
export const XL_EQUAL = 3;
export const XL_NOT_EQUAL = 4;
export const XL_GREATER = 5;
export const XL_LESS = 6;
export const XL_GREATER_EQUAL = 7;
export const XL_LESS_EQUAL = 8;

/** XlDVType */
export const XL_VALIDATE_INPUT_ONLY = 0;
export const XL_VALIDATE_WHOLE = 1;
export const XL_VALIDATE_DECIMAL = 2;
export const XL_VALIDATE_LIST = 3;
export const XL_VALIDATE_DATE = 4;
export const XL_VALIDATE_TIME = 5;
export const XL_VALIDATE_TEXT_LENGTH = 6;
export const XL_VALIDATE_CUSTOM = 7;

/** xlValidAlertStop */
export const XL_VALID_ALERT_STOP = 1;

export const CF_EVIDENCE =
  "WPS JSA FormatConditions member-probe + mock tests (desktop COM FormatConditions.Add parity); not real device sideload";

export const DV_EVIDENCE =
  "WPS JSA Validation member-probe + mock tests (desktop COM Validation.Add/Delete parity); not real device sideload";

const CF_OP_TO_COM: Record<CellValueOperator, number> = {
  between: XL_BETWEEN,
  notBetween: XL_NOT_BETWEEN,
  equalTo: XL_EQUAL,
  notEqualTo: XL_NOT_EQUAL,
  greaterThan: XL_GREATER,
  lessThan: XL_LESS,
  greaterThanOrEqualTo: XL_GREATER_EQUAL,
  lessThanOrEqualTo: XL_LESS_EQUAL,
};

const CF_COM_TO_OP: Record<number, CellValueOperator> = {
  [XL_BETWEEN]: "between",
  [XL_NOT_BETWEEN]: "notBetween",
  [XL_EQUAL]: "equalTo",
  [XL_NOT_EQUAL]: "notEqualTo",
  [XL_GREATER]: "greaterThan",
  [XL_LESS]: "lessThan",
  [XL_GREATER_EQUAL]: "greaterThanOrEqualTo",
  [XL_LESS_EQUAL]: "lessThanOrEqualTo",
};

const DV_TYPE_TO_COM: Record<DataValidationType, number> = {
  wholeNumber: XL_VALIDATE_WHOLE,
  decimal: XL_VALIDATE_DECIMAL,
  list: XL_VALIDATE_LIST,
  date: XL_VALIDATE_DATE,
  time: XL_VALIDATE_TIME,
  textLength: XL_VALIDATE_TEXT_LENGTH,
  custom: XL_VALIDATE_CUSTOM,
};

const DV_COM_TO_TYPE: Record<number, DataValidationType> = {
  [XL_VALIDATE_WHOLE]: "wholeNumber",
  [XL_VALIDATE_DECIMAL]: "decimal",
  [XL_VALIDATE_LIST]: "list",
  [XL_VALIDATE_DATE]: "date",
  [XL_VALIDATE_TIME]: "time",
  [XL_VALIDATE_TEXT_LENGTH]: "textLength",
  [XL_VALIDATE_CUSTOM]: "custom",
};

const DV_HOST_LABEL: Record<DataValidationType, string> = {
  wholeNumber: "WholeNumber",
  decimal: "Decimal",
  list: "List",
  date: "Date",
  time: "Time",
  textLength: "TextLength",
  custom: "Custom",
};

export function mapCfOperatorToCom(op: CellValueOperator): number {
  return CF_OP_TO_COM[op];
}

export function unmapCfOperatorFromCom(raw: unknown): CellValueOperator | undefined {
  const n = toInt(raw);
  if (n == null) return undefined;
  return CF_COM_TO_OP[n];
}

export function mapDvTypeToCom(type: DataValidationType): number {
  return DV_TYPE_TO_COM[type];
}

export function unmapDvTypeFromCom(raw: unknown): DataValidationType | undefined {
  const n = toInt(raw);
  if (n == null) return undefined;
  return DV_COM_TO_TYPE[n];
}

export function mapDvOperatorToCom(op: DataValidationOperator): number {
  return CF_OP_TO_COM[op];
}

export function unmapDvOperatorFromCom(raw: unknown): DataValidationOperator | undefined {
  return unmapCfOperatorFromCom(raw);
}

export function dvHostTypeLabel(type: DataValidationType): string {
  return DV_HOST_LABEL[type];
}

export function classifyCfComType(raw: unknown): {
  kind: ConditionalFormatListKind;
  hostType: string;
  supported: boolean;
  limitations?: string[];
} {
  const n = toInt(raw);
  if (n === XL_CELL_VALUE) {
    return { kind: "cellValue", hostType: "CellValue", supported: true };
  }
  if (n === XL_EXPRESSION) {
    return { kind: "custom", hostType: "Custom", supported: true };
  }
  const hostType =
    n == null ? `Unknown(${String(raw)})` : `FormatConditionType(${n})`;
  return {
    kind: "unsupported",
    hostType,
    supported: false,
    limitations: [`WPS FormatCondition type ${hostType} is not add/verify-capable in this add-in`],
  };
}

export function toInt(raw: unknown): number | null {
  if (typeof raw === "number" && Number.isFinite(raw)) return Math.trunc(raw);
  if (typeof raw === "string" && /^-?\d+$/.test(raw.trim())) {
    return Number.parseInt(raw.trim(), 10);
  }
  return null;
}

export function asBool(raw: unknown): boolean | undefined {
  if (typeof raw === "boolean") return raw;
  if (raw === -1 || raw === 1) return true;
  if (raw === 0) return false;
  return undefined;
}

export function formulaText(raw: unknown): string {
  if (raw == null) return "";
  return String(raw);
}
