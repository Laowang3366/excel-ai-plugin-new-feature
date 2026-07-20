/**
 * Phase53.3: enum readback is case-insensitive exact only (no trim / space collapse).
 */
import { afterEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import {
  assertOfficialDvHostType,
  unmapAlertStyle,
} from "../shared/host/officeJsValidationAlerts";
import { unmapDvOperator } from "../shared/host/officeJsValidationMapping";
import { installValidationExcel } from "./fakes/officeJsValidationFake";

describe("phase53.3 exact enum readback", () => {
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
    delete (globalThis as { Office?: unknown }).Office;
  });

  it("assertOfficialDvHostType rejects spaced/padded/hyphenated variants", () => {
    expect(assertOfficialDvHostType("WholeNumber")).toBe("WholeNumber");
    expect(assertOfficialDvHostType("wholenumber")).toBe("WholeNumber");
    expect(assertOfficialDvHostType("TextLength")).toBe("TextLength");
    expect(assertOfficialDvHostType("MixedCriteria")).toBe("MixedCriteria");

    for (const bad of [
      "Whole Number",
      " WholeNumber",
      "WholeNumber ",
      "Whole-Number",
      "Mixed Criteria",
      "Text Length",
      " None",
      "None ",
    ]) {
      expect(() => assertOfficialDvHostType(bad), bad).toThrow(/unknown|not string/i);
    }
  });

  it("unmapDvOperator rejects spaced/padded/hyphenated host operators", () => {
    expect(unmapDvOperator("GreaterThan")).toBe("greaterThan");
    expect(unmapDvOperator("greaterthanorequalto")).toBe("greaterThanOrEqualTo");
    expect(unmapDvOperator("Between")).toBe("between");

    for (const bad of [
      "Greater Than",
      "GreaterThan ",
      " GreaterThan",
      "Greater-Than",
      "Greater Than Or Equal To",
      "Not Equal To",
    ]) {
      expect(unmapDvOperator(bad), bad).toBeUndefined();
    }
  });

  it("unmapAlertStyle rejects padded/aliased styles", () => {
    expect(unmapAlertStyle("Stop")).toBe("stop");
    expect(unmapAlertStyle("WARNING")).toBe("warning");
    expect(unmapAlertStyle("Information")).toBe("information");

    for (const bad of [" Stop ", "Stop ", " Stop", "Warning ", "Information-Alert", "stop "]) {
      expect(unmapAlertStyle(bad), JSON.stringify(bad)).toBeUndefined();
    }
  });

  it("write post-read poison bad type/operator/style fails without echoing request", async () => {
    // Bad style after write — host coercion, side effects may exist; must not ok with request echo.
    installValidationExcel({
      tamperDvReadback: { errorAlertStyle: " Stop " },
    });
    const badStyle = await new OfficeJsAdapter().writeDataValidation({
      sheetName: "Sheet1",
      range: "A1",
      rule: { type: "wholeNumber", operator: "equalTo", formula1: "1" },
      errorAlert: { showAlert: true, style: "stop", title: "T", message: "M" },
    });
    expect(badStyle.ok).toBe(false);
    if (!badStyle.ok) {
      expect(badStyle.unsupported).not.toBe(true);
      expect(badStyle.reason).not.toMatch(/"style":"stop"/);
      expect(badStyle.reason).toMatch(/errorAlert\.style|unknown/i);
    }

    installValidationExcel({
      tamperDvReadback: { operator: "Greater Than" },
    });
    const badOp = await new OfficeJsAdapter().writeDataValidation({
      sheetName: "Sheet1",
      range: "A2",
      rule: { type: "wholeNumber", operator: "greaterThan", formula1: "0" },
    });
    expect(badOp.ok).toBe(false);
    if (!badOp.ok) {
      expect(badOp.unsupported).not.toBe(true);
      // must not report success with requested operator
      expect(JSON.stringify(badOp)).not.toMatch(/"operator":"greaterThan"/);
    }

    // poisonSurface applies on every load including pre-read → zero writes
    const ctrl = installValidationExcel({ poisonSurface: { type: "Whole Number" } });
    ctrl.resetDvWriteCounts();
    const badType = await new OfficeJsAdapter().writeDataValidation({
      sheetName: "Sheet1",
      range: "A3",
      rule: { type: "list", listValues: ["a"] },
    });
    expect(badType.ok).toBe(false);
    if (!badType.ok) {
      expect(badType.unsupported).not.toBe(true);
      expect(badType.reason).toMatch(/type|unknown/i);
    }
    expect(ctrl.getDvWriteCounts()).toEqual({
      rule: 0,
      ignoreBlanks: 0,
      errorAlert: 0,
      prompt: 0,
    });
  });
});
