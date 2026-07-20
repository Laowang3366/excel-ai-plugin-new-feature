/**
 * Conditional format (ExcelApi 1.6) + data validation (ExcelApi 1.8).
 * - Requirement-set precheck → typed unsupported
 * - Post-precheck member/sync/business errors → ordinary failed
 * - CF list: one batch load/sync (O(1) in rule count)
 * - CF add: load id/type before first sync; verify rule/colors after second sync
 * - DV write: full rule + ignoreBlanks + errorAlert/prompt match after readback
 * - List source: string inline or Excel.Range proxy (never String(object))
 */
import type { ExcelRequestContext } from "./officeJsRuntime";
import { withExcel } from "./officeJsRuntime";
import { cfRuleFieldsMatch } from "./officeJsValidationCompare";
import {
  classifyCfHostType,
  mapCfOperatorToHost,
  unmapCfOperator,
} from "./officeJsValidationMapping";
import {
  requireExcelApiForCf,
  requireExcelApiForDv,
} from "./officeJsValidationRequirements";
import {
  assertLoadedDvSurface,
  assignErrorAlert,
  assignPrompt,
  DV_FULL_LOAD_PROPS,
  mergeErrorAlertForWrite,
  mergePromptForWrite,
} from "./officeJsValidationAlerts";
import { planDvRuleWrite } from "./officeJsValidationPlan";
import {
  assertDvCleared,
  assertDvWriteMatches,
  parseDvRule,
  toDvInfo,
} from "./officeJsValidationReadback";
import type {
  ConditionalFormatInfo,
  ConditionalFormatRule,
  DataValidationInfo,
  DataValidationWriteInput,
  HostResult,
} from "./types";

const CELL_VALUE = "CellValue";
const CUSTOM = "Custom";

export async function officeJsListConditionalFormats(
  sheetName: string,
  rangeAddress: string,
): Promise<HostResult<ConditionalFormatInfo[]>> {
  const pre = requireExcelApiForCf("conditionalFormat.list");
  if (pre) return pre;
  return withExcel("conditionalFormat.list", async (context: ExcelRequestContext) => {
    const range = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.load("address");
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

function queueCfRule(
  cf: {
    cellValue?: {
      rule: { formula1: string; formula2?: string; operator: string };
      format: { fill: { color: string }; font: { color: string } };
      load?: (props: string) => void;
    };
    custom?: {
      rule: { formula: string; load?: (props: string) => void };
      format: { fill: { color: string }; font: { color: string } };
    };
  },
  rule: ConditionalFormatRule,
): void {
  if (rule.kind === "custom") {
    if (!cf.custom) throw new Error("Custom conditional format not available");
    cf.custom.rule.formula = rule.formula!;
    if (rule.fillColor) cf.custom.format.fill.color = rule.fillColor;
    if (rule.fontColor) cf.custom.format.font.color = rule.fontColor;
    return;
  }
  if (!cf.cellValue) throw new Error("CellValue conditional format not available");
  cf.cellValue.rule = {
    operator: mapCfOperatorToHost(rule.operator!),
    formula1: rule.formula1!,
    ...(rule.formula2 != null ? { formula2: rule.formula2 } : {}),
  };
  if (rule.fillColor) cf.cellValue.format.fill.color = rule.fillColor;
  if (rule.fontColor) cf.cellValue.format.font.color = rule.fontColor;
}

function queueCfDetailLoad(cf: {
  cellValue?: {
    rule: { formula1: string; formula2?: string; operator: string };
    format: { fill: { color: string; load?: (p: string) => void }; font: { color: string; load?: (p: string) => void } };
    load?: (props: string) => void;
  };
  custom?: {
    rule: { formula: string; load?: (props: string) => void };
    format: { fill: { color: string; load?: (p: string) => void }; font: { color: string; load?: (p: string) => void } };
  };
  type: string;
}): void {
  if (cf.type === CELL_VALUE && cf.cellValue) {
    cf.cellValue.load?.("rule");
    cf.cellValue.format.fill.load?.("color");
    cf.cellValue.format.font.load?.("color");
    return;
  }
  if (cf.type === CUSTOM && cf.custom) {
    cf.custom.rule.load?.("formula");
    cf.custom.format.fill.load?.("color");
    cf.custom.format.font.load?.("color");
  }
}

function readCfHostFields(cf: {
  type: string;
  cellValue?: {
    rule: { formula1: string; formula2?: string; operator: string };
    format: { fill: { color: string }; font: { color: string } };
  };
  custom?: {
    rule: { formula: string };
    format: { fill: { color: string }; font: { color: string } };
  };
}): {
  operator?: string;
  formula1?: string;
  formula2?: string;
  formula?: string;
  fillColor?: string;
  fontColor?: string;
} {
  if (cf.type === CELL_VALUE && cf.cellValue) {
    return {
      operator: cf.cellValue.rule.operator,
      formula1: cf.cellValue.rule.formula1,
      formula2: cf.cellValue.rule.formula2,
      fillColor: cf.cellValue.format.fill.color,
      fontColor: cf.cellValue.format.font.color,
    };
  }
  if (cf.type === CUSTOM && cf.custom) {
    return {
      formula: cf.custom.rule.formula,
      fillColor: cf.custom.format.fill.color,
      fontColor: cf.custom.format.font.color,
    };
  }
  return {};
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
    // Must load add() proxy before first sync — collection load does not cover this proxy.
    cf.load("id,type");
    queueCfRule(cf, input.rule);
    await context.sync();
    queueCfDetailLoad(cf);
    await context.sync();
    const classified = classifyCfHostType(cf.type);
    if (classified.kind !== input.rule.kind) {
      throw new Error(
        `conditional format kind mismatch after add: expected ${input.rule.kind}, got ${classified.hostType}`,
      );
    }
    const hostFields = readCfHostFields(cf);
    if (!cfRuleFieldsMatch(input.rule, hostFields, input.sheetName)) {
      throw new Error(
        `conditional format rule/color mismatch after add: host=${JSON.stringify(hostFields)}`,
      );
    }
    return {
      id: cf.id,
      sheetName: input.sheetName,
      range: range.address,
      kind: classified.kind,
      hostType: classified.hostType,
      supported: classified.supported,
      summary: `${classified.hostType}:${cf.id}`,
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
    if (range.conditionalFormats.items.some((c) => c.id === id)) {
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
    range.dataValidation.load(DV_FULL_LOAD_PROPS);
    await context.sync();
    assertLoadedDvSurface(range.dataValidation);
    const parsed = await parseDvRule(range.dataValidation, context);
    return toDvInfo(sheetName, range.address, parsed);
  });
}

export async function officeJsWriteDataValidation(
  input: DataValidationWriteInput,
): Promise<HostResult<DataValidationInfo>> {
  const pre = requireExcelApiForDv("dataValidation.write");
  if (pre) return pre;
  return withExcel("dataValidation.write", async (context: ExcelRequestContext) => {
    const range = context.workbook.worksheets.getItem(input.sheetName).getRange(input.range);
    const dv = range.dataValidation;
    if (typeof dv.load !== "function") {
      throw new Error("dataValidation.load is missing");
    }

    // A) Pre-read only — no business writes yet (ClientObject PropertyNotLoaded otherwise).
    range.load("address");
    dv.load(DV_FULL_LOAD_PROPS);
    await context.sync();

    // B) Validate complete official surface after first sync; zero write side effects if bad.
    const loaded = assertLoadedDvSurface(dv);

    // C) Plan ALL business payloads before any assignment (pure planner + merge).
    const plannedRule = planDvRuleWrite(context, input.sheetName, input.rule);
    const plannedIgnore = input.rule.allowBlank !== false;
    const plannedError =
      input.errorAlert !== undefined
        ? mergeErrorAlertForWrite(loaded.errorAlert, input.errorAlert)
        : undefined;
    const plannedPrompt =
      input.prompt !== undefined
        ? mergePromptForWrite(loaded.prompt, input.prompt)
        : undefined;

    // D) Single assignment pass after full plan.
    dv.rule = plannedRule;
    dv.ignoreBlanks = plannedIgnore;
    if (plannedError !== undefined) assignErrorAlert(dv, plannedError);
    if (plannedPrompt !== undefined) assignPrompt(dv, plannedPrompt);

    await context.sync();
    dv.load(DV_FULL_LOAD_PROPS);
    await context.sync();
    // Re-validate surface + parse for honest host snapshot.
    assertLoadedDvSurface(dv);
    const parsed = await parseDvRule(dv, context);
    assertDvWriteMatches(
      input.rule,
      parsed,
      input.sheetName,
      input.errorAlert,
      input.prompt,
    );
    return toDvInfo(input.sheetName, range.address, parsed);
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
    range.dataValidation.load(DV_FULL_LOAD_PROPS);
    await context.sync();
    assertLoadedDvSurface(range.dataValidation);
    const parsed = await parseDvRule(range.dataValidation, context);
    assertDvCleared(parsed);
    // Range.address already includes Sheet!A1 — do not re-prefix sheetName.
    return { cleared: range.address };
  });
}

// re-export unmap for tests that import from this module historically
export { unmapCfOperator };
