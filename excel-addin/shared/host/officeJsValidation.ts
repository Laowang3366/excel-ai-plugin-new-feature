/**
 * Conditional format (ExcelApi 1.6) + data validation (ExcelApi 1.8).
 * - Requirement-set precheck → typed unsupported
 * - Post-precheck member/sync/business errors → ordinary failed
 * - CF list: one batch load/sync (O(1) in rule count)
 * - Write paths: queue → sync → load/readback → success only after host confirms
 */
import { normalizeSameSheetA1Range, parseChartSourceRange } from "./officeJsChartSource";
import type { ExcelDataValidation, ExcelRequestContext } from "./officeJsRuntime";
import { withExcel } from "./officeJsRuntime";
import {
  classifyCfHostType,
  classifyDvHostType,
  classifyListSource,
  COMPARE_DV_TYPES,
  isBetweenOp,
  mapCfOperatorToHost,
  mapDvOperatorToHost,
  unmapDvOperator,
} from "./officeJsValidationMapping";
import {
  requireExcelApiForCf,
  requireExcelApiForDv,
} from "./officeJsValidationRequirements";
import type {
  ConditionalFormatInfo,
  ConditionalFormatRule,
  DataValidationInfo,
  DataValidationRule,
  DataValidationType,
  HostResult,
} from "./types";

const CELL_VALUE = "CellValue";
const CUSTOM = "Custom";

type CompareBag = {
  formula1?: string | number;
  formula2?: string | number;
  operator?: string;
};

function parseCompareRule(
  type: Exclude<DataValidationType, "list" | "custom">,
  bag: CompareBag | undefined,
  allowBlank: boolean | undefined,
): DataValidationRule {
  return {
    type,
    operator: unmapDvOperator(bag?.operator),
    formula1: bag?.formula1 != null ? String(bag.formula1) : undefined,
    formula2: bag?.formula2 != null ? String(bag.formula2) : undefined,
    allowBlank,
  };
}

export function parseDvRule(dv: ExcelDataValidation): {
  rule: DataValidationRule | null;
  hostType: string;
  supported: boolean;
  listSourceKind?: "inline" | "range" | null;
  limitations?: string[];
} {
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
  const raw = (dv.rule ?? {}) as Record<string, CompareBag | { source?: unknown; formula?: string }>;
  const allowBlank = dv.ignoreBlanks;
  if (classified.type === "list") {
    const list = raw.list as { source?: unknown } | undefined;
    const sourceText =
      typeof list?.source === "string"
        ? list.source
        : list?.source != null
          ? String(list.source)
          : "";
    const classifiedSource = classifyListSource(sourceText);
    if (classifiedSource.kind === "range") {
      return {
        rule: {
          type: "list",
          formula1: classifiedSource.formula1,
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
        listValues: classifiedSource.listValues,
        // Keep raw source only as non-formula echo when useful; do not invent formula1 for inline.
        allowBlank,
      },
      hostType: "List",
      supported: true,
      listSourceKind: "inline",
    };
  }
  if (classified.type === "custom") {
    const custom = raw.custom as { formula?: string } | undefined;
    return {
      rule: {
        type: "custom",
        formula1: custom?.formula != null ? String(custom.formula) : undefined,
        allowBlank,
      },
      hostType: "Custom",
      supported: true,
    };
  }
  const bagKey = classified.type as "wholeNumber" | "decimal" | "date" | "time" | "textLength";
  const bag = raw[bagKey] as CompareBag | undefined;
  return {
    rule: parseCompareRule(classified.type, bag, allowBlank),
    hostType: classified.hostType,
    supported: true,
  };
}

function resolveListSourceRange(
  context: ExcelRequestContext,
  ownerSheetName: string,
  formula1: string,
): { range: ReturnType<ExcelRequestContext["workbook"]["worksheets"]["getItem"]> extends never ? never : object; display: string } {
  const raw = formula1.trim().replace(/^=/, "");
  // Same-sheet bare A1 or qualified Sheet!A1 / 'Sheet'!A1
  if (!raw.includes("!")) {
    const bare = normalizeSameSheetA1Range(ownerSheetName, raw, "formula1", "dataValidation");
    const sheet = context.workbook.worksheets.getItem(ownerSheetName);
    return { range: sheet.getRange(bare), display: bare };
  }
  const parsed = parseChartSourceRange(ownerSheetName, raw);
  const sheet = context.workbook.worksheets.getItem(parsed.sourceSheetName);
  return { range: sheet.getRange(parsed.bareA1), display: parsed.displaySourceRange };
}

export async function officeJsListConditionalFormats(
  sheetName: string,
  rangeAddress: string,
): Promise<HostResult<ConditionalFormatInfo[]>> {
  const pre = requireExcelApiForCf("conditionalFormat.list");
  if (pre) return pre;
  return withExcel("conditionalFormat.list", async (context: ExcelRequestContext) => {
    const range = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.load("address");
    // Single batch: items/id,type — no per-rule sync.
    range.conditionalFormats.load("items/id,items/type");
    await context.sync();
    const result: ConditionalFormatInfo[] = [];
    for (const item of range.conditionalFormats.items) {
      const classified = classifyCfHostType(item.type);
      result.push({
        id: item.id,
        sheetName,
        range: range.address,
        kind: classified.kind,
        hostType: classified.hostType,
        supported: classified.supported,
        summary: `${classified.hostType}:${item.id}`,
        limitations: classified.limitations,
      });
    }
    return result;
  });
}

export async function officeJsAddConditionalFormat(input: {
  sheetName: string;
  range: string;
  rule: ConditionalFormatRule;
}): Promise<HostResult<ConditionalFormatInfo>> {
  const pre = requireExcelApiForCf("conditionalFormat.add");
  if (pre) return pre;
  return withExcel("conditionalFormat.add", async (context: ExcelRequestContext) => {
    const range = context.workbook.worksheets.getItem(input.sheetName).getRange(input.range);
    range.load("address");
    const type = input.rule.kind === "custom" ? CUSTOM : CELL_VALUE;
    const cf = range.conditionalFormats.add(type);
    if (input.rule.kind === "custom") {
      if (!cf.custom) throw new Error("Custom conditional format not available");
      // ClientObject: queue formula property (not whole rule object replace).
      cf.custom.rule.formula = input.rule.formula!;
      if (input.rule.fillColor) cf.custom.format.fill.color = input.rule.fillColor;
      if (input.rule.fontColor) cf.custom.format.font.color = input.rule.fontColor;
    } else {
      if (!cf.cellValue) throw new Error("CellValue conditional format not available");
      // Plain rule data: whole-object assign queues correctly.
      cf.cellValue.rule = {
        operator: mapCfOperatorToHost(input.rule.operator!),
        formula1: input.rule.formula1!,
        ...(input.rule.formula2 != null ? { formula2: input.rule.formula2 } : {}),
      };
      if (input.rule.fillColor) cf.cellValue.format.fill.color = input.rule.fillColor;
      if (input.rule.fontColor) cf.cellValue.format.font.color = input.rule.fontColor;
    }
    await context.sync();
    // Readback from host collection
    range.conditionalFormats.load("items/id,items/type");
    await context.sync();
    const found = range.conditionalFormats.items.find((item) => item.id === cf.id);
    if (!found) throw new Error("conditional format missing after add readback");
    const classified = classifyCfHostType(found.type);
    if (classified.kind !== input.rule.kind) {
      throw new Error(
        `conditional format kind mismatch after add: expected ${input.rule.kind}, got ${classified.hostType}`,
      );
    }
    return {
      id: found.id,
      sheetName: input.sheetName,
      range: range.address,
      kind: classified.kind,
      hostType: classified.hostType,
      supported: classified.supported,
      summary: `${classified.hostType}:${found.id}`,
    };
  });
}

export async function officeJsDeleteConditionalFormat(
  sheetName: string,
  rangeAddress: string,
  id: string,
): Promise<HostResult<{ deleted: string }>> {
  const pre = requireExcelApiForCf("conditionalFormat.delete");
  if (pre) return pre;
  return withExcel("conditionalFormat.delete", async (context: ExcelRequestContext) => {
    const range = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    const item = range.conditionalFormats.getItem(id);
    item.delete();
    await context.sync();
    range.conditionalFormats.load("items/id");
    await context.sync();
    if (range.conditionalFormats.items.some((cf) => cf.id === id)) {
      throw new Error(`conditional format still present after delete: ${id}`);
    }
    return { deleted: id };
  });
}

export async function officeJsReadDataValidation(
  sheetName: string,
  rangeAddress: string,
): Promise<HostResult<DataValidationInfo>> {
  const pre = requireExcelApiForDv("dataValidation.read");
  if (pre) return pre;
  return withExcel("dataValidation.read", async (context: ExcelRequestContext) => {
    const range = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.load("address");
    range.dataValidation.load("type,rule,ignoreBlanks");
    await context.sync();
    const parsed = parseDvRule(range.dataValidation);
    return {
      sheetName,
      range: range.address,
      rule: parsed.rule,
      hostType: parsed.hostType,
      supported: parsed.supported,
      listSourceKind: parsed.listSourceKind ?? null,
      limitations: parsed.limitations,
    };
  });
}

function applyCompareDv(
  dv: ExcelDataValidation,
  rule: DataValidationRule,
): void {
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

export async function officeJsWriteDataValidation(input: {
  sheetName: string;
  range: string;
  rule: DataValidationRule;
}): Promise<HostResult<DataValidationInfo>> {
  const pre = requireExcelApiForDv("dataValidation.write");
  if (pre) return pre;
  return withExcel("dataValidation.write", async (context: ExcelRequestContext) => {
    const range = context.workbook.worksheets.getItem(input.sheetName).getRange(input.range);
    range.load("address");
    const dv = range.dataValidation;
    if (input.rule.type === "list") {
      if (input.rule.listValues && input.rule.listValues.length > 0) {
        if (input.rule.listValues.some((v) => v.includes(","))) {
          throw new Error("listValues items must not contain commas; use a range source instead");
        }
        const source = input.rule.listValues.join(",");
        dv.rule = { list: { inCellDropDown: true, source } };
      } else if (input.rule.formula1) {
        const resolved = resolveListSourceRange(context, input.sheetName, input.rule.formula1);
        // Official contract: source is string inline list OR Excel.Range proxy.
        dv.rule = { list: { inCellDropDown: true, source: resolved.range as unknown as string } };
      } else {
        throw new Error("list validation requires listValues or formula1 range source");
      }
    } else if (input.rule.type === "custom") {
      if (!input.rule.formula1) throw new Error("custom requires formula1");
      dv.rule = { custom: { formula: input.rule.formula1 } } as ExcelDataValidation["rule"];
    } else {
      applyCompareDv(dv, input.rule);
    }
    dv.ignoreBlanks = input.rule.allowBlank !== false;
    await context.sync();
    dv.load("type,rule,ignoreBlanks");
    await context.sync();
    const parsed = parseDvRule(dv);
    if (!parsed.supported || !parsed.rule) {
      throw new Error(
        `data validation readback not supported after write: ${parsed.hostType}`,
      );
    }
    if (parsed.rule.type !== input.rule.type) {
      throw new Error(
        `data validation type mismatch after write: expected ${input.rule.type}, got ${parsed.rule.type}`,
      );
    }
    return {
      sheetName: input.sheetName,
      range: range.address,
      rule: parsed.rule,
      hostType: parsed.hostType,
      supported: true,
      listSourceKind: parsed.listSourceKind ?? null,
      limitations: parsed.limitations,
    };
  });
}

export async function officeJsClearDataValidation(
  sheetName: string,
  rangeAddress: string,
): Promise<HostResult<{ cleared: string }>> {
  const pre = requireExcelApiForDv("dataValidation.clear");
  if (pre) return pre;
  return withExcel("dataValidation.clear", async (context: ExcelRequestContext) => {
    const range = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.load("address");
    range.dataValidation.clear();
    await context.sync();
    range.dataValidation.load("type,rule,ignoreBlanks");
    await context.sync();
    const parsed = parseDvRule(range.dataValidation);
    if (parsed.rule != null && parsed.supported) {
      throw new Error("data validation still present after clear readback");
    }
    return { cleared: `${sheetName}!${range.address}` };
  });
}
