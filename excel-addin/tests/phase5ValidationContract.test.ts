import { describe, expect, it } from "vitest";
import {
  classifyListSource,
  formulasSemanticallyEqual,
  tryParseSimpleA1Parts,
} from "../shared/host/officeJsValidationCompare";
import {
  mapCfOperatorToHost,
  mapDvOperatorToHost,
  unmapCfOperator,
  unmapDvOperator,
  classifyCfHostType,
} from "../shared/host/officeJsValidationMapping";

describe("phase5 validation contract (maps + sources)", () => {
  it("maps GTE/LTE and NotEqualTo correctly; CF has no OrEqualTo suffix", () => {
    expect(mapCfOperatorToHost("greaterThanOrEqualTo")).toBe("GreaterThanOrEqual");
    expect(mapCfOperatorToHost("lessThanOrEqualTo")).toBe("LessThanOrEqual");
    expect(mapCfOperatorToHost("notEqualTo")).toBe("NotEqualTo");
    expect(unmapCfOperator("NotEqualTo")).toBe("notEqualTo");
    expect(unmapCfOperator("NotEqual")).toBeUndefined();
    expect(mapDvOperatorToHost("greaterThanOrEqualTo")).toBe("GreaterThanOrEqualTo");
    expect(mapDvOperatorToHost("lessThanOrEqualTo")).toBe("LessThanOrEqualTo");
    expect(unmapDvOperator("GreaterThanOrEqualTo")).toBe("greaterThanOrEqualTo");
    expect(unmapCfOperator("GreaterThanOrEqualTo")).toBeUndefined();
    expect(unmapDvOperator("GreaterThanOrEqual")).toBeUndefined();
  });

  it("classifies CF host types without disguising ContainsText as cellValue", () => {
    expect(classifyCfHostType("ContainsText")).toMatchObject({
      kind: "unsupported",
      hostType: "ContainsText",
      supported: false,
    });
    expect(classifyCfHostType("CellValue").kind).toBe("cellValue");
  });

  it("single-token host list sources are inline (not null, not range)", () => {
    for (const raw of ["Yes", "1", "x"]) {
      const c = classifyListSource(raw);
      expect(c.kind, raw).toBe("inline");
      expect(c.listValues, raw).toEqual([raw]);
      expect(c.lossy, raw).toBeFalsy();
    }
  });

  it("only lossless same-workbook A1 is kind=range; never for names/functions/external/3D", () => {
    const good = classifyListSource("=Sheet1!$A$1:$A$3");
    expect(good.kind).toBe("range");
    expect(good.lossy).toBeFalsy();
    expect(classifyListSource("A1:A3").kind).toBe("range");
    expect(classifyListSource("工作表1!A1").kind).toBe("range");
    expect(classifyListSource("'Q1-Final'!A1").kind).toBe("range");
    expect(classifyListSource("'North, South'!A1").kind).toBe("range");
    expect(classifyListSource("'Q1 (Final)'!A1").kind).toBe("range");
    expect(classifyListSource("'O''Brien'!A1").kind).toBe("range");

    const bangInline = classifyListSource("Yes!,No");
    expect(bangInline.kind).toBe("inline");
    expect(bangInline.listValues).toEqual(["Yes!", "No"]);

    // A0 is not a valid A1 range row; under list classification it is a plain inline token.
    expect(classifyListSource("A0").kind).toBe("inline");
    expect(classifyListSource("A0").listValues).toEqual(["A0"]);
    expect(classifyListSource("A00").kind).toBe("inline");
    expect(classifyListSource("A00").listValues).toEqual(["A00"]);

    for (const bad of [
      "=MyList",
      '=INDIRECT("A1")',
      "=[Book.xlsx]Sheet1!A1",
      "=Sheet1:Sheet3!A1",
      "=Table1[Col]",
      "=Sheet1!A1,Sheet1!B1",
      "Sheet 1!A1", // unquoted space is illegal range form
    ]) {
      const c = classifyListSource(bad);
      expect(c.kind, bad).toBeNull();
      expect(c.lossy, bad).toBe(true);
    }

    const lossy = classifyListSource("A,,B");
    expect(lossy.kind).toBe("inline");
    expect(lossy.lossy).toBe(true);
  });

  it("tryParseSimpleA1Parts rejects illegal sheets and A0", () => {
    expect(tryParseSimpleA1Parts("A1")).toEqual({ sheet: null, a1: "A1" });
    expect(tryParseSimpleA1Parts("'Sheet 1'!$A$1")?.a1).toBe("A1");
    expect(tryParseSimpleA1Parts("Sheet 1!A1")).toBeNull();
    expect(tryParseSimpleA1Parts("A0")).toBeNull();
    expect(tryParseSimpleA1Parts("工作表1!B2")?.sheet).toBe("工作表1");
  });

  it("formula equality is owner-scoped; case/literals sensitive", () => {
    expect(formulasSemanticallyEqual("1", "=1")).toBe(true);
    expect(formulasSemanticallyEqual("$A$1", "A1")).toBe(true);
    expect(formulasSemanticallyEqual("Sheet1!$A$1", "Sheet1!A1")).toBe(true);
    expect(formulasSemanticallyEqual("A1", "Sheet1!A1")).toBe(false);
    expect(formulasSemanticallyEqual("A1", "Sheet1!A1", "Sheet1")).toBe(true);
    expect(formulasSemanticallyEqual("A1", "Sheet1!A1", "Sheet2")).toBe(false);
    expect(
      formulasSemanticallyEqual("'Sheet 1'!$A$1", "'Sheet 1'!A1", "Sheet 1"),
    ).toBe(true);
    expect(formulasSemanticallyEqual('=EXACT("A","a")', '=EXACT("a","a")')).toBe(false);
    expect(formulasSemanticallyEqual('="$100"', '="100"')).toBe(false);
    expect(formulasSemanticallyEqual("=SUM(A1)", "=sum(a1)")).toBe(false);
  });
});
