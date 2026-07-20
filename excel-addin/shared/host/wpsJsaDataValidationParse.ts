/**
 * WPS Validation read/snapshot helpers.
 */
import {
  classifyListSource,
  hostHasExtraFormula2,
} from "./officeJsValidationCompare";
import { isBetweenOp } from "./officeJsValidationMapping";
import type {
  DataValidationListSourceKind,
  DataValidationRule,
} from "./types";
import type { WpsValidation } from "./wpsJsaRuntime";
import {
  XL_VALIDATE_CUSTOM,
  XL_VALIDATE_INPUT_ONLY,
  XL_VALIDATE_LIST,
  XL_VALID_ALERT_STOP,
  asBool,
  dvHostTypeLabel,
  formulaText,
  toInt,
  unmapDvOperatorFromCom,
  unmapDvTypeFromCom,
} from "./wpsJsaValidationConstants";

export type ValidationSnapshot = {
  type: number;
  operator?: number;
  formula1?: string;
  formula2?: string;
  ignoreBlank?: boolean;
  inCellDropdown?: boolean;
};

export type ParsedWpsValidation = {
  rule: DataValidationRule | null;
  hostType: string;
  supported: boolean;
  listSourceKind?: DataValidationListSourceKind | null;
  limitations?: string[];
};

export function trySnapshot(v: WpsValidation): ValidationSnapshot | null {
  try {
    const type = toInt(v.Type);
    if (type == null || type === XL_VALIDATE_INPUT_ONLY) return null;
    const op = toInt(v.Operator);
    return {
      type,
      operator: op ?? undefined,
      formula1: formulaText(v.Formula1) || undefined,
      formula2: formulaText(v.Formula2) || undefined,
      ignoreBlank: asBool(v.IgnoreBlank),
      inCellDropdown: asBool(v.InCellDropdown),
    };
  } catch {
    return null;
  }
}

export function restoreSnapshot(v: WpsValidation, snap: ValidationSnapshot): void {
  if (typeof v.Add !== "function") {
    throw new Error("Validation.Add unavailable during restore");
  }
  if (snap.type === XL_VALIDATE_LIST || snap.type === XL_VALIDATE_CUSTOM) {
    v.Add(snap.type, XL_VALID_ALERT_STOP, undefined, snap.formula1 ?? "");
  } else {
    v.Add(
      snap.type,
      XL_VALID_ALERT_STOP,
      snap.operator ?? undefined,
      snap.formula1 ?? "",
      snap.formula2,
    );
  }
  if (snap.ignoreBlank != null && v.IgnoreBlank !== undefined) {
    v.IgnoreBlank = snap.ignoreBlank;
  }
  if (snap.inCellDropdown != null && v.InCellDropdown !== undefined) {
    v.InCellDropdown = snap.inCellDropdown;
  }
}

export function parseHostValidation(
  v: WpsValidation,
  ownerSheetName: string,
): ParsedWpsValidation {
  let typeRaw: unknown;
  try {
    typeRaw = v.Type;
  } catch {
    return { rule: null, hostType: "None", supported: false };
  }
  const typeNum = toInt(typeRaw);
  if (typeNum == null) {
    return {
      rule: null,
      hostType: `Unknown(${String(typeRaw)})`,
      supported: false,
      limitations: [`Unknown Validation.Type ${String(typeRaw)}`],
    };
  }
  if (typeNum === XL_VALIDATE_INPUT_ONLY) {
    return { rule: null, hostType: "None", supported: false };
  }
  const type = unmapDvTypeFromCom(typeNum);
  if (!type) {
    return {
      rule: null,
      hostType: `ValidationType(${typeNum})`,
      supported: false,
      limitations: [`Unsupported Validation.Type ${typeNum}`],
    };
  }
  const hostType = dvHostTypeLabel(type);
  const allowBlank = asBool(v.IgnoreBlank) !== false;
  const formula1 = formulaText(v.Formula1);
  const formula2 = formulaText(v.Formula2);

  if (type === "list") {
    const classified = classifyListSource(formula1);
    if (classified.kind === "inline" && !classified.lossy) {
      return {
        rule: { type: "list", listValues: classified.listValues, allowBlank },
        hostType,
        supported: true,
        listSourceKind: "inline",
      };
    }
    if (classified.kind === "range" && !classified.lossy) {
      return {
        rule: {
          type: "list",
          formula1: classified.formula1 ?? formula1,
          allowBlank,
        },
        hostType,
        supported: true,
        listSourceKind: "range",
      };
    }
    return {
      rule: null,
      hostType,
      supported: false,
      listSourceKind: classified.kind,
      limitations: classified.limitations ?? ["list source is not a writable shape"],
    };
  }

  if (type === "custom") {
    if (hostHasExtraFormula2(undefined, formula2)) {
      return {
        rule: null,
        hostType,
        supported: false,
        limitations: ["custom Validation has unexpected non-empty formula2 from host"],
      };
    }
    if (!formula1) {
      return {
        rule: null,
        hostType,
        supported: false,
        limitations: ["custom Validation missing formula"],
      };
    }
    return {
      rule: { type: "custom", formula1, allowBlank },
      hostType,
      supported: true,
    };
  }

  const operator = unmapDvOperatorFromCom(v.Operator);
  if (!operator || !formula1) {
    return {
      rule: null,
      hostType,
      supported: false,
      limitations: ["compare Validation missing operator/formula1"],
    };
  }
  if (hostHasExtraFormula2(operator, formula2) && !isBetweenOp(operator)) {
    return {
      rule: null,
      hostType,
      supported: false,
      limitations: ["non-between Validation has unexpected non-empty formula2"],
    };
  }
  void ownerSheetName;
  return {
    rule: {
      type,
      operator,
      formula1,
      allowBlank,
      ...(isBetweenOp(operator) ? { formula2: formula2 || undefined } : {}),
    },
    hostType,
    supported: true,
  };
}
