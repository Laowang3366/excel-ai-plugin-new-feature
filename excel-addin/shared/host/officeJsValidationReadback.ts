/**
 * CF/DV parse + list-source materialization + write match helpers.
 * Keeps Office.js validation orchestration under the line limit.
 */
import { normalizeSameSheetA1Range, parseChartSourceRange } from "./officeJsChartSource";
import type { ExcelDataValidation, ExcelRange, ExcelRequestContext } from "./officeJsRuntime";
import {
  classifyListSource,
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
import type {
  DataValidationInfo,
  DataValidationRule,
  DataValidationType,
} from "./types";

export type ExcelClientObjectRange = {
  load: (props: string) => void;
  address: string;
};

export type ExcelPlainAddressRange = {
  address: string;
};

/** Real Office.js Range proxy: has load(); do NOT read address before load+sync. */
export function isClientObjectRange(value: unknown): value is ExcelClientObjectRange {
  return (
    typeof value === "object" &&
    value !== null &&
    typeof (value as { load?: unknown }).load === "function"
  );
}

/** Already-materialized plain { address } (tests / non-ClientObject). */
export function isPlainAddressRange(value: unknown): value is ExcelPlainAddressRange {
  if (typeof value !== "object" || value === null) return false;
  if (typeof (value as { load?: unknown }).load === "function") return false;
  return typeof (value as { address?: unknown }).address === "string";
}

export function resolveListSourceRange(
  context: ExcelRequestContext,
  ownerSheetName: string,
  formula1: string,
): ExcelRange {
  const raw = formula1.trim().replace(/^=/, "");
  if (!raw.includes("!")) {
    const bare = normalizeSameSheetA1Range(ownerSheetName, raw, "formula1", "dataValidation");
    return context.workbook.worksheets.getItem(ownerSheetName).getRange(bare);
  }
  const parsed = parseChartSourceRange(ownerSheetName, raw);
  return context.workbook.worksheets
    .getItem(parsed.sourceSheetName)
    .getRange(parsed.bareA1);
}

type CompareBag = {
  formula1?: string | number;
  formula2?: string | number;
  operator?: string;
};

/**
 * Resolve list.source that may be string or Excel.Range ClientObject.
 * When Range-like, load address if needed (caller must sync when load queued).
 */
export async function materializeListSource(
  source: unknown,
  context: ExcelRequestContext,
): Promise<{
  kind: "inline" | "range" | null;
  formula1?: string;
  listValues?: string[];
  lossy?: boolean;
  limitations?: string[];
}> {
  if (source == null) {
    return {
      kind: null,
      limitations: ["list source is empty"],
    };
  }
  if (isClientObjectRange(source)) {
    // Always load+sync before reading address (PropertyNotLoaded otherwise).
    source.load("address");
    await context.sync();
    let address = "";
    try {
      address = String(source.address ?? "").trim();
    } catch (err) {
      return {
        kind: null,
        limitations: [
          `list Range source address PropertyNotLoaded after load: ${
            err instanceof Error ? err.message : String(err)
          }`,
        ],
      };
    }
    if (!address) {
      return {
        kind: null,
        limitations: ["list Range source address unavailable after load"],
      };
    }
    return { kind: "range", formula1: address };
  }
  if (isPlainAddressRange(source)) {
    const address = source.address.trim();
    if (!address) {
      return {
        kind: null,
        limitations: ["list Range source address empty"],
      };
    }
    return { kind: "range", formula1: address };
  }
  if (typeof source !== "string") {
    return {
      kind: null,
      limitations: [
        `list source is not string or Range (got ${typeof source}); not coerced via String(object)`,
      ],
    };
  }
  const classified = classifyListSource(source);
  if (classified.lossy) {
    return {
      kind: classified.kind,
      formula1: classified.formula1,
      listValues: classified.listValues,
      lossy: true,
      limitations:
        classified.limitations ??
        [`list source is lossy/unparseable: ${classified.raw ?? source}`],
    };
  }
  if (classified.kind === "range") {
    return { kind: "range", formula1: classified.formula1 };
  }
  return { kind: "inline", listValues: classified.listValues };
}

export async function parseDvRule(
  dv: ExcelDataValidation,
  context: ExcelRequestContext,
): Promise<{
  rule: DataValidationRule | null;
  hostType: string;
  supported: boolean;
  listSourceKind?: "inline" | "range" | null;
  limitations?: string[];
}> {
  const classified = classifyDvHostType(dv.type);
  if (classified.mixedState || (!classified.writable && classified.type === null)) {
    return {
      rule: null,
      hostType: classified.hostType,
      supported: false,
      limitations: classified.limitations,
    };
  }
  if (classified.type === null) {
    return {
      rule: null,
      hostType: classified.hostType,
      supported: false,
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
    };
  }

  if (classified.type === "custom") {
    const custom = raw.custom as { formula?: string } | undefined;
    const formula =
      custom?.formula != null && String(custom.formula).trim() !== ""
        ? String(custom.formula)
        : undefined;
    if (!formula) {
      return {
        rule: null,
        hostType: "Custom",
        supported: false,
        limitations: ["custom DataValidation missing formula — not a writable rule"],
      };
    }
    return {
      rule: { type: "custom", formula1: formula, allowBlank },
      hostType: "Custom",
      supported: true,
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
        limitations: [`${classified.hostType} ${operator} missing formula2`],
      };
    }
    return {
      rule: { type: classified.type, operator, formula1, formula2, allowBlank },
      hostType: classified.hostType,
      supported: true,
    };
  }
  if (hostHasExtraFormula2(operator, bag?.formula2)) {
    return {
      rule: null,
      hostType: classified.hostType,
      supported: false,
      limitations: [
        `${classified.hostType} ${operator} has unexpected non-empty formula2 from host`,
      ],
    };
  }
  return {
    rule: { type: classified.type, operator, formula1, allowBlank },
    hostType: classified.hostType,
    supported: true,
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
    limitations: parsed.limitations,
  };
}
