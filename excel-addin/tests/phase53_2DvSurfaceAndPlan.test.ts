/**
 * Phase53.2: complete surface assert + pure planDvRuleWrite before host assignment.
 * All illegal cases call OfficeJsAdapter directly (not only ToolExecutor).
 */
import { afterEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { planDvRuleWrite } from "../shared/host/officeJsValidationPlan";
import { assertOfficialDvHostType } from "../shared/host/officeJsValidationAlerts";
import { installValidationExcel } from "./fakes/officeJsValidationFake";

function zeroWrites(ctrl: ReturnType<typeof installValidationExcel>) {
  expect(ctrl.getDvWriteCounts()).toEqual({
    rule: 0,
    ignoreBlanks: 0,
    errorAlert: 0,
    prompt: 0,
  });
}

describe("phase53.2 DV surface + pure rule plan", () => {
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
    delete (globalThis as { Office?: unknown }).Office;
  });

  it("assertOfficialDvHostType accepts official tokens only", () => {
    expect(assertOfficialDvHostType("None")).toBe("None");
    expect(assertOfficialDvHostType("wholeNumber")).toBe("WholeNumber");
    expect(assertOfficialDvHostType("MixedCriteria")).toBe("MixedCriteria");
    expect(() => assertOfficialDvHostType(undefined)).toThrow(/null|undefined|not string/i);
    expect(() => assertOfficialDvHostType(1)).toThrow(/not string/);
    expect(() => assertOfficialDvHostType("NotAType")).toThrow(/unknown/);
    expect(() => assertOfficialDvHostType("Whole-Number")).toThrow(/unknown/);
  });

  it("write: between missing formula2 fails before any host assignment", async () => {
    const ctrl = installValidationExcel();
    ctrl.resetDvWriteCounts();
    const r = await new OfficeJsAdapter().writeDataValidation({
      sheetName: "Sheet1",
      range: "A1",
      rule: {
        type: "wholeNumber",
        operator: "between",
        formula1: "1",
        // missing formula2
      },
    });
    expect(r.ok).toBe(false);
    if (!r.ok) {
      expect(r.unsupported).not.toBe(true);
      expect(r.reason).toMatch(/formula2/);
    }
    zeroWrites(ctrl);
  });

  it("write: non-between extra formula2 fails before assignment", async () => {
    const ctrl = installValidationExcel();
    ctrl.resetDvWriteCounts();
    const r = await new OfficeJsAdapter().writeDataValidation({
      sheetName: "Sheet1",
      range: "A1",
      rule: {
        type: "decimal",
        operator: "greaterThan",
        formula1: "0",
        formula2: "9",
      },
    });
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.reason).toMatch(/must not include formula2/);
    zeroWrites(ctrl);
  });

  it("write: list listValues+formula1 XOR fails before assignment", async () => {
    const ctrl = installValidationExcel();
    ctrl.resetDvWriteCounts();
    const r = await new OfficeJsAdapter().writeDataValidation({
      sheetName: "Sheet1",
      range: "A1",
      rule: {
        type: "list",
        listValues: ["A"],
        formula1: "B1:B2",
      },
    });
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.reason).toMatch(/mutually exclusive|listValues/);
    zeroWrites(ctrl);
  });

  it("write: list with operator/formula2 fails before assignment", async () => {
    const ctrl = installValidationExcel();
    ctrl.resetDvWriteCounts();
    const withOp = await new OfficeJsAdapter().writeDataValidation({
      sheetName: "Sheet1",
      range: "A1",
      rule: { type: "list", listValues: ["A"], operator: "equalTo" as never },
    });
    expect(withOp.ok).toBe(false);
    zeroWrites(ctrl);

    ctrl.resetDvWriteCounts();
    const withF2 = await new OfficeJsAdapter().writeDataValidation({
      sheetName: "Sheet1",
      range: "A1",
      rule: { type: "list", listValues: ["A"], formula2: "x" },
    });
    expect(withF2.ok).toBe(false);
    zeroWrites(ctrl);
  });

  it("write: custom with operator/formula2/listValues fails before assignment", async () => {
    const ctrl = installValidationExcel();
    for (const rule of [
      { type: "custom" as const, formula1: "=A1>0", operator: "equalTo" as never },
      { type: "custom" as const, formula1: "=A1>0", formula2: "1" },
      { type: "custom" as const, formula1: "=A1>0", listValues: ["x"] },
    ]) {
      ctrl.resetDvWriteCounts();
      const r = await new OfficeJsAdapter().writeDataValidation({
        sheetName: "Sheet1",
        range: "A1",
        rule,
      });
      expect(r.ok, JSON.stringify(rule)).toBe(false);
      zeroWrites(ctrl);
    }
  });

  it("write: unknown type fails before assignment", async () => {
    const ctrl = installValidationExcel();
    ctrl.resetDvWriteCounts();
    const r = await new OfficeJsAdapter().writeDataValidation({
      sheetName: "Sheet1",
      range: "A1",
      rule: { type: "regex" as never, formula1: "x" },
    });
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.reason).toMatch(/type|list\|wholeNumber/);
    zeroWrites(ctrl);
  });

  it("write pre-read: bad type/ignoreBlanks/null rule → ordinary failed, zero writes", async () => {
    for (const poison of [
      { type: 123 as unknown },
      { type: "NotARealType" },
      { ignoreBlanks: "yes" as unknown },
      { rule: null as unknown },
      { rule: "nope" as unknown },
    ]) {
      const ctrl = installValidationExcel({ poisonSurface: poison });
      ctrl.resetDvWriteCounts();
      const r = await new OfficeJsAdapter().writeDataValidation({
        sheetName: "Sheet1",
        range: "A1",
        rule: { type: "list", listValues: ["ok"] },
      });
      expect(r.ok, JSON.stringify(poison)).toBe(false);
      if (!r.ok) {
        expect(r.unsupported).not.toBe(true);
      }
      zeroWrites(ctrl);
    }
  });

  it("read: bad surface ordinary failed (not malformed ok data)", async () => {
    for (const poison of [
      { type: 99 as unknown },
      { ignoreBlanks: 1 as unknown },
      { rule: null as unknown },
    ]) {
      installValidationExcel({ poisonSurface: poison });
      const r = await new OfficeJsAdapter().readDataValidation("Sheet1", "A1");
      expect(r.ok, JSON.stringify(poison)).toBe(false);
      if (!r.ok) expect(r.unsupported).not.toBe(true);
    }
  });

  it("read: official None surface still succeeds", async () => {
    installValidationExcel();
    const r = await new OfficeJsAdapter().readDataValidation("Sheet1", "Z9");
    expect(r.ok).toBe(true);
    if (r.ok) {
      expect(r.data.hostType).toBe("None");
      expect(r.data.rule).toBeNull();
    }
  });

  it("requirement 1.8 false: typed unsupported, Excel.run 0", async () => {
    const ctrl = installValidationExcel({ excelApi18: false });
    ctrl.resetSyncCount();
    ctrl.resetDvWriteCounts();
    const r = await new OfficeJsAdapter().writeDataValidation({
      sheetName: "Sheet1",
      range: "A1",
      rule: { type: "list", listValues: ["a"] },
    });
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.unsupported).toBe(true);
    expect(ctrl.getSyncCount()).toBe(0);
    zeroWrites(ctrl);
  });

  it("planDvRuleWrite is pure and strict for between formula2", () => {
    const ctx = {
      workbook: {
        worksheets: {
          getItem: () => ({
            getRange: (a: string) => ({ address: a, load: () => undefined }),
          }),
        },
      },
      sync: async () => undefined,
    } as never;
    expect(() =>
      planDvRuleWrite(ctx, "Sheet1", {
        type: "wholeNumber",
        operator: "between",
        formula1: "1",
      }),
    ).toThrow(/formula2/);
    const planned = planDvRuleWrite(ctx, "Sheet1", {
      type: "wholeNumber",
      operator: "between",
      formula1: "1",
      formula2: "10",
    });
    expect(planned.wholeNumber?.formula2).toBe("10");
  });
});
