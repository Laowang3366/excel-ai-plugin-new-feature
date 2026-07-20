/**
 * WPS JSA conditional formats via Range.FormatConditions (COM-style).
 * Public IDs are stable 1-based collection indexes as decimal strings ("1","2",…).
 * After delete, host renumbers remaining items — callers must re-list.
 */
import {
  cfRuleFieldsMatch,
} from "./officeJsValidationCompare";
import { isBetweenOp, mapCfOperatorToHost } from "./officeJsValidationMapping";
import type {
  ConditionalFormatInfo,
  ConditionalFormatRule,
  HostResult,
} from "./types";
import { fail, ok, unsupported } from "./types";
import { hexFromOleColor, oleColorFromHex } from "./wpsJsaFormat";
import {
  getSheet,
  requireWorkbook,
  type WpsFormatCondition,
  type WpsFormatConditions,
  type WpsRange,
} from "./wpsJsaRuntime";
import {
  CF_EVIDENCE as EVIDENCE,
  XL_CELL_VALUE,
  XL_EXPRESSION,
  classifyCfComType,
  formulaText,
  mapCfOperatorToCom,
  unmapCfOperatorFromCom,
} from "./wpsJsaValidationConstants";

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function resolveRange(
  capability: string,
  sheetName: string,
  address: string,
): HostResult<{ range: WpsRange; sheetName: string; address: string }> {
  const workbookResult = requireWorkbook(capability);
  if (!workbookResult.ok) return workbookResult;
  const sheet = getSheet(workbookResult.data, sheetName);
  if (!sheet?.Range) {
    return unsupported(
      capability,
      "wps-jsa",
      `Sheet "${sheetName}" or Range API missing`,
      EVIDENCE,
    );
  }
  try {
    const range = sheet.Range(address);
    return ok({
      range,
      sheetName,
      address: String(range.Address ?? `${sheetName}!${address}`),
    });
  } catch (error) {
    return fail(capability, "wps-jsa", messageOf(error), EVIDENCE);
  }
}

function requireFormatConditions(
  capability: string,
  range: WpsRange,
): HostResult<WpsFormatConditions> {
  const fc = range.FormatConditions;
  if (!fc || typeof fc !== "object") {
    return unsupported(
      capability,
      "wps-jsa",
      "Range.FormatConditions is unavailable",
      EVIDENCE,
    );
  }
  if (typeof fc.Count !== "number" || typeof fc.Item !== "function") {
    return unsupported(
      capability,
      "wps-jsa",
      "FormatConditions.Count/Item is unavailable",
      EVIDENCE,
    );
  }
  return ok(fc);
}

function readConditionFields(item: WpsFormatCondition): {
  operator?: string;
  formula1?: string;
  formula2?: string;
  formula?: string;
  fillColor?: string;
  fontColor?: string;
} {
  const classified = classifyCfComType(item.Type);
  const fillColor = item.Interior ? hexFromOleColor(item.Interior.Color) ?? undefined : undefined;
  const fontColor = item.Font ? hexFromOleColor(item.Font.Color) ?? undefined : undefined;
  if (classified.kind === "custom") {
    return {
      formula: formulaText(item.Formula1),
      fillColor: fillColor ?? undefined,
      fontColor: fontColor ?? undefined,
    };
  }
  const op = unmapCfOperatorFromCom(item.Operator);
  return {
    operator: op,
    formula1: formulaText(item.Formula1),
    formula2: formulaText(item.Formula2) || undefined,
    fillColor: fillColor ?? undefined,
    fontColor: fontColor ?? undefined,
  };
}

function fingerprint(item: WpsFormatCondition): string {
  const f = readConditionFields(item);
  return [
    String(item.Type ?? ""),
    f.operator ?? "",
    f.formula1 ?? f.formula ?? "",
    f.formula2 ?? "",
    f.fillColor ?? "",
    f.fontColor ?? "",
  ].join("|");
}

function listFingerprints(fc: WpsFormatConditions): string[] {
  const count = Math.max(0, Math.trunc(fc.Count ?? 0));
  const out: string[] = [];
  for (let i = 1; i <= count; i += 1) {
    const item = fc.Item?.(i);
    out.push(item ? fingerprint(item) : `missing:${i}`);
  }
  return out;
}

function applyColors(item: WpsFormatCondition, rule: ConditionalFormatRule): void {
  if (rule.fillColor) {
    if (!item.Interior || item.Interior.Color === undefined) {
      throw new Error("FormatCondition.Interior.Color is unavailable");
    }
    const ole = oleColorFromHex(rule.fillColor);
    if (ole == null) throw new Error(`invalid fillColor ${rule.fillColor}`);
    item.Interior.Color = ole;
  }
  if (rule.fontColor) {
    if (!item.Font || item.Font.Color === undefined) {
      throw new Error("FormatCondition.Font.Color is unavailable");
    }
    const ole = oleColorFromHex(rule.fontColor);
    if (ole == null) throw new Error(`invalid fontColor ${rule.fontColor}`);
    item.Font.Color = ole;
  }
}

function toInfo(
  id: string,
  sheetName: string,
  rangeAddress: string,
  item: WpsFormatCondition,
): ConditionalFormatInfo {
  const classified = classifyCfComType(item.Type);
  return {
    id,
    sheetName,
    range: rangeAddress,
    kind: classified.kind,
    hostType: classified.hostType,
    supported: classified.supported,
    summary: `${classified.hostType}:${id}`,
    limitations: classified.limitations,
  };
}

export async function wpsListConditionalFormats(
  sheetName: string,
  rangeAddress: string,
): Promise<HostResult<ConditionalFormatInfo[]>> {
  const resolved = resolveRange("conditionalFormat.list", sheetName, rangeAddress);
  if (!resolved.ok) return resolved;
  const fcResult = requireFormatConditions("conditionalFormat.list", resolved.data.range);
  if (!fcResult.ok) return fcResult;
  try {
    const fc = fcResult.data;
    const count = Math.max(0, Math.trunc(fc.Count ?? 0));
    const out: ConditionalFormatInfo[] = [];
    for (let i = 1; i <= count; i += 1) {
      const item = fc.Item?.(i);
      if (!item) {
        return fail(
          "conditionalFormat.list",
          "wps-jsa",
          `FormatConditions.Item(${i}) missing`,
          EVIDENCE,
        );
      }
      out.push(toInfo(String(i), sheetName, resolved.data.address, item));
    }
    return ok(out);
  } catch (error) {
    return fail("conditionalFormat.list", "wps-jsa", messageOf(error), EVIDENCE);
  }
}

export async function wpsAddConditionalFormat(input: {
  sheetName: string;
  range: string;
  rule: ConditionalFormatRule;
}): Promise<HostResult<ConditionalFormatInfo>> {
  const capability = "conditionalFormat.add";
  const resolved = resolveRange(capability, input.sheetName, input.range);
  if (!resolved.ok) return resolved;
  const fcResult = requireFormatConditions(capability, resolved.data.range);
  if (!fcResult.ok) return fcResult;
  const fc = fcResult.data;
  if (typeof fc.Add !== "function") {
    return unsupported(capability, "wps-jsa", "FormatConditions.Add is unavailable", EVIDENCE);
  }
  const rule = input.rule;
  if (rule.kind === "cellValue") {
    if (!rule.operator || rule.formula1 == null || rule.formula1 === "") {
      return fail(capability, "wps-jsa", "cellValue requires operator and formula1", EVIDENCE);
    }
    if (isBetweenOp(rule.operator) && (rule.formula2 == null || rule.formula2 === "")) {
      return fail(capability, "wps-jsa", "between/notBetween requires formula2", EVIDENCE);
    }
    if (!isBetweenOp(rule.operator) && rule.formula2 != null && String(rule.formula2).trim() !== "") {
      return fail(
        capability,
        "wps-jsa",
        "formula2 is only allowed for between/notBetween",
        EVIDENCE,
      );
    }
  } else if (rule.kind === "custom") {
    if (rule.formula == null || rule.formula === "") {
      return fail(capability, "wps-jsa", "custom requires formula", EVIDENCE);
    }
  } else {
    return fail(capability, "wps-jsa", `unsupported CF kind`, EVIDENCE);
  }

  try {
    const before = Math.max(0, Math.trunc(fc.Count ?? 0));
    let created: WpsFormatCondition;
    if (rule.kind === "custom") {
      created = fc.Add(XL_EXPRESSION, undefined, rule.formula);
    } else {
      const op = mapCfOperatorToCom(rule.operator!);
      created = isBetweenOp(rule.operator)
        ? fc.Add(XL_CELL_VALUE, op, rule.formula1!, rule.formula2!)
        : fc.Add(XL_CELL_VALUE, op, rule.formula1!);
    }
    if (!created) {
      return fail(capability, "wps-jsa", "FormatConditions.Add returned empty", EVIDENCE);
    }
    applyColors(created, rule);

    const after = Math.max(0, Math.trunc(fc.Count ?? 0));
    if (after !== before + 1) {
      return fail(
        capability,
        "wps-jsa",
        `FormatConditions count after add expected ${before + 1}, got ${after}`,
        EVIDENCE,
      );
    }
    const id = String(after);
    const item = fc.Item?.(after);
    if (!item) {
      return fail(capability, "wps-jsa", `FormatConditions.Item(${after}) missing after add`, EVIDENCE);
    }
    const hostFields = readConditionFields(item);
    if (!matchCfRuleLocal(rule, hostFields, input.sheetName)) {
      return fail(
        capability,
        "wps-jsa",
        `conditional format readback mismatch after add for id ${id}`,
        EVIDENCE,
      );
    }
    return ok(toInfo(id, input.sheetName, resolved.data.address, item));
  } catch (error) {
    return fail(capability, "wps-jsa", messageOf(error), EVIDENCE);
  }
}

function matchCfRuleLocal(
  expected: ConditionalFormatRule,
  host: {
    operator?: string;
    formula1?: string;
    formula2?: string;
    formula?: string;
    fillColor?: string;
    fontColor?: string;
  },
  ownerSheetName: string,
): boolean {
  if (expected.kind === "cellValue") {
    if (host.operator !== expected.operator) return false;
    return cfRuleFieldsMatch(
      expected,
      {
        ...host,
        operator: mapCfOperatorToHost(expected.operator!),
      },
      ownerSheetName,
    );
  }
  return cfRuleFieldsMatch(expected, host, ownerSheetName);
}

export async function wpsDeleteConditionalFormat(
  sheetName: string,
  rangeAddress: string,
  id: string,
): Promise<HostResult<{ deleted: string }>> {
  const capability = "conditionalFormat.delete";
  const resolved = resolveRange(capability, sheetName, rangeAddress);
  if (!resolved.ok) return resolved;
  const fcResult = requireFormatConditions(capability, resolved.data.range);
  if (!fcResult.ok) return fcResult;
  const fc = fcResult.data;
  if (!/^\d+$/.test(id.trim())) {
    return fail(capability, "wps-jsa", `invalid FormatConditions index id: ${id}`, EVIDENCE);
  }
  const index = Number.parseInt(id.trim(), 10);
  if (index < 1) {
    return fail(capability, "wps-jsa", `invalid FormatConditions index id: ${id}`, EVIDENCE);
  }
  try {
    const before = Math.max(0, Math.trunc(fc.Count ?? 0));
    if (index > before) {
      return fail(
        capability,
        "wps-jsa",
        `conditional format id ${id} not found (count=${before})`,
        EVIDENCE,
      );
    }
    const target = fc.Item?.(index);
    if (!target || typeof target.Delete !== "function") {
      return unsupported(
        capability,
        "wps-jsa",
        "FormatCondition.Delete is unavailable",
        EVIDENCE,
      );
    }
    const beforePrints = listFingerprints(fc);
    const removed = beforePrints[index - 1]!;
    target.Delete();
    const after = Math.max(0, Math.trunc(fc.Count ?? 0));
    if (after !== before - 1) {
      return fail(
        capability,
        "wps-jsa",
        `FormatConditions count after delete expected ${before - 1}, got ${after}`,
        EVIDENCE,
      );
    }
    const afterPrints = listFingerprints(fc);
    const expected = [...beforePrints.slice(0, index - 1), ...beforePrints.slice(index)];
    if (
      afterPrints.length !== expected.length ||
      afterPrints.some((fp, i) => fp !== expected[i])
    ) {
      return fail(
        capability,
        "wps-jsa",
        `conditional format delete left unexpected remaining set (removed ${removed})`,
        EVIDENCE,
      );
    }
    return ok({ deleted: String(index) });
  } catch (error) {
    return fail(capability, "wps-jsa", messageOf(error), EVIDENCE);
  }
}
