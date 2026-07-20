/**
 * CF/DV parse + write match helpers (list source in officeJsValidationListSource).
 */
import type { ExcelDataValidation, ExcelRequestContext } from "./officeJsRuntime";
import {
  dvRulesMatch,
  hostHasExtraFormula2,
} from "./officeJsValidationCompare";
import {
  classifyDvHostType,
  COMPARE_DV_TYPES,
  isBetweenOp,
  mapDvOperatorToHost,
  unmapDvOperator,
} from "./officeJsValidationMapping";
import {
  errorAlertMatches,
  parseErrorAlertFromHost,
  parsePromptFromHost,
  promptMatches,
} from "./officeJsValidationAlerts";
import { materializeListSource } from "./officeJsValidationListSource";
import type {
  DataValidationErrorAlert,
  DataValidationInfo,
  DataValidationPrompt,
  DataValidationRule,
  DataValidationType,
} from "./types";

// Re-export for existing officeJsValidation imports
export { resolveListSourceRange } from "./officeJsValidationListSource";

type CompareBag = {
  formula1?: string | number;
  formula2?: string | number;
  operator?: string;
};

export async function parseDvRule(
  dv: ExcelDataValidation,
  context: ExcelRequestContext,
): Promise<{
  rule: DataValidationRule | null;
  hostType: string;
  supported: boolean;
  listSourceKind?: "inline" | "range" | null;
  errorAlert?: DataValidationErrorAlert | null;
  prompt?: DataValidationPrompt | null;
  limitations?: string[];
}> {
  const eaParsed = parseErrorAlertFromHost(dv.errorAlert);
  if (eaParsed.error) {
    throw new Error(eaParsed.error);
  }
  const prParsed = parsePromptFromHost(dv.prompt);
  if (prParsed.error) {
    throw new Error(prParsed.error);
  }
  const alertSnap = {
    errorAlert: eaParsed.value,
    prompt: prParsed.value,
  };
  const classified = classifyDvHostType(dv.type);
  if (classified.mixedState || (!classified.writable && classified.type === null)) {
    return {
      rule: null,
      hostType: classified.hostType,
      supported: false,
      errorAlert: alertSnap.errorAlert,
      prompt: alertSnap.prompt,
      limitations: classified.limitations,
    };
  }
  if (classified.type === null) {
    return {
      rule: null,
      hostType: classified.hostType,
      supported: false,
      errorAlert: alertSnap.errorAlert,
      prompt: alertSnap.prompt,
      limitations: classified.limitations,
    };
  }
  const raw = (dv.rule ?? {}) as Record<
    string,
    CompareBag | { source?: unknown; formula?: string }
  >;
  const allowBlank = dv.ignoreBlanks;

  if (classified.type === "list") {
    const list = raw.list as { source?: unknown } | undefined;
    const sourceInfo = await materializeListSource(list?.source, context);
    if (sourceInfo.kind === null || sourceInfo.lossy) {
      return {
        rule: null,
        hostType: "List",
        supported: false,
        listSourceKind: sourceInfo.kind,
        errorAlert: alertSnap.errorAlert,
        prompt: alertSnap.prompt,
        limitations: sourceInfo.limitations,
      };
    }
    if (sourceInfo.kind === "range") {
      return {
        rule: {
          type: "list",
          formula1: sourceInfo.formula1,
          allowBlank,
        },
        hostType: "List",
        supported: true,
        listSourceKind: "range",
        errorAlert: alertSnap.errorAlert,
        prompt: alertSnap.prompt,
      };
    }
    return {
      rule: {
        type: "list",
        listValues: sourceInfo.listValues,
        allowBlank,
      },
      hostType: "List",
      supported: true,
      listSourceKind: "inline",
      errorAlert: alertSnap.errorAlert,
      prompt: alertSnap.prompt,
    };
  }

  if (classified.type === "custom") {
    const custom = raw.custom as { formula?: string; formula2?: string | number } | undefined;
    if (hostHasExtraFormula2(undefined, custom?.formula2)) {
      return {
        rule: null,
        hostType: "Custom",
        supported: false,
        errorAlert: alertSnap.errorAlert,
        prompt: alertSnap.prompt,
        limitations: [
          "custom DataValidation has unexpected non-empty formula2 from host",
        ],
      };
    }
    const formula =
      custom?.formula != null && String(custom.formula).trim() !== ""
        ? String(custom.formula)
        : undefined;
    if (!formula) {
      return {
        rule: null,
        hostType: "Custom",
        supported: false,
        errorAlert: alertSnap.errorAlert,
        prompt: alertSnap.prompt,
        limitations: ["custom DataValidation missing formula — not a writable rule"],
      };
    }
    return {
      rule: { type: "custom", formula1: formula, allowBlank },
      hostType: "Custom",
      supported: true,
      errorAlert: alertSnap.errorAlert,
      prompt: alertSnap.prompt,
    };
  }

  const bagKey = classified.type as "wholeNumber" | "decimal" | "date" | "time" | "textLength";
  const bag = raw[bagKey] as CompareBag | undefined;
  const operator = unmapDvOperator(bag?.operator);
  const formula1 =
    bag?.formula1 != null && String(bag.formula1).trim() !== ""
      ? String(bag.formula1)
      : undefined;
  if (!operator || !formula1) {
    return {
      rule: null,
      hostType: classified.hostType,
      supported: false,
      errorAlert: alertSnap.errorAlert,
      prompt: alertSnap.prompt,
      limitations: [
        `${classified.hostType} missing recognized operator or formula1 — not supported:true with synthetic rule`,
      ],
    };
  }
  if (isBetweenOp(operator)) {
    const formula2 =
      bag?.formula2 != null && String(bag.formula2).trim() !== ""
        ? String(bag.formula2)
        : undefined;
    if (!formula2) {
      return {
        rule: null,
        hostType: classified.hostType,
        supported: false,
        errorAlert: alertSnap.errorAlert,
        prompt: alertSnap.prompt,
        limitations: [`${classified.hostType} ${operator} missing formula2`],
      };
    }
    return {
      rule: { type: classified.type, operator, formula1, formula2, allowBlank },
      hostType: classified.hostType,
      supported: true,
      errorAlert: alertSnap.errorAlert,
      prompt: alertSnap.prompt,
    };
  }
  if (hostHasExtraFormula2(operator, bag?.formula2)) {
    return {
      rule: null,
      hostType: classified.hostType,
      supported: false,
      errorAlert: alertSnap.errorAlert,
      prompt: alertSnap.prompt,
      limitations: [
        `${classified.hostType} ${operator} has unexpected non-empty formula2 from host`,
      ],
    };
  }
  return {
    rule: { type: classified.type, operator, formula1, allowBlank },
    hostType: classified.hostType,
    supported: true,
    errorAlert: alertSnap.errorAlert,
    prompt: alertSnap.prompt,
  };
}

export function applyCompareDv(dv: ExcelDataValidation, rule: DataValidationRule): void {
  const type = rule.type as Exclude<DataValidationType, "list" | "custom">;
  if (!COMPARE_DV_TYPES.includes(type)) throw new Error(`not a compare DV type: ${type}`);
  if (!rule.operator || !rule.formula1) throw new Error(`${type} requires operator and formula1`);
  const bag = {
    formula1: rule.formula1,
    operator: mapDvOperatorToHost(rule.operator),
    ...(isBetweenOp(rule.operator) && rule.formula2 != null
      ? { formula2: rule.formula2 }
      : {}),
  };
  dv.rule = { [type]: bag } as ExcelDataValidation["rule"];
}

export function assertDvWriteMatches(
  expected: DataValidationRule,
  parsed: Awaited<ReturnType<typeof parseDvRule>>,
  ownerSheetName: string,
  expectedErrorAlert?: DataValidationErrorAlert,
  expectedPrompt?: DataValidationPrompt,
): void {
  if (!parsed.supported || !parsed.rule) {
    throw new Error(
      `data validation readback not supported after write: ${parsed.hostType}`,
    );
  }
  if (!dvRulesMatch(expected, parsed.rule, parsed.listSourceKind, ownerSheetName)) {
    throw new Error(
      `data validation rule mismatch after write: expected ${JSON.stringify(expected)}, got ${JSON.stringify(parsed.rule)} listKind=${parsed.listSourceKind}`,
    );
  }
  if (expectedErrorAlert !== undefined) {
    if (!errorAlertMatches(expectedErrorAlert, parsed.errorAlert)) {
      throw new Error(
        `data validation errorAlert mismatch after write: expected ${JSON.stringify(expectedErrorAlert)}, got ${JSON.stringify(parsed.errorAlert)}`,
      );
    }
  }
  if (expectedPrompt !== undefined) {
    if (!promptMatches(expectedPrompt, parsed.prompt)) {
      throw new Error(
        `data validation prompt mismatch after write: expected ${JSON.stringify(expectedPrompt)}, got ${JSON.stringify(parsed.prompt)}`,
      );
    }
  }
}

export function assertDvCleared(parsed: Awaited<ReturnType<typeof parseDvRule>>): void {
  const hostType = parsed.hostType;
  if (hostType !== "None" || parsed.rule != null || parsed.supported) {
    throw new Error(
      `data validation clear readback not None: hostType=${hostType}, rule=${parsed.rule != null}, supported=${parsed.supported}`,
    );
  }
}

export function toDvInfo(
  sheetName: string,
  rangeAddress: string,
  parsed: Awaited<ReturnType<typeof parseDvRule>>,
): DataValidationInfo {
  return {
    sheetName,
    range: rangeAddress,
    rule: parsed.rule,
    hostType: parsed.hostType,
    supported: parsed.supported,
    listSourceKind: parsed.listSourceKind ?? null,
    errorAlert: parsed.errorAlert ?? null,
    prompt: parsed.prompt ?? null,
    limitations: parsed.limitations,
  };
}
