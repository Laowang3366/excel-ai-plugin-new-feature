import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { installValidationExcel } from "./fakes/officeJsValidationFake";

describe("phase5 Office.js CF/DV", () => {
  let gates: ReturnType<typeof installValidationExcel>;

  beforeEach(() => {
    gates = installValidationExcel();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  it("adds cellValue CF, lists, deletes via conditionalFormats API", async () => {
    const adapter = new OfficeJsAdapter();
    const added = await adapter.addConditionalFormat({
      sheetName: "Sheet1",
      range: "A1:A10",
      rule: {
        kind: "cellValue",
        operator: "greaterThan",
        formula1: "10",
        fillColor: "#FF0000",
      },
    });
    expect(added.ok).toBe(true);
    if (!added.ok) return;
    expect(gates.getCfCount("Sheet1", "A1:A10")).toBe(1);
    const listed = await adapter.listConditionalFormats("Sheet1", "A1:A10");
    expect(listed.ok).toBe(true);
    if (listed.ok) {
      expect(listed.data.some((item) => item.id === added.data.id)).toBe(true);
      expect(listed.data).toHaveLength(1);
    }
    const deleted = await adapter.deleteConditionalFormat("Sheet1", "A1:A10", added.data.id);
    expect(deleted.ok).toBe(true);
    expect(gates.getCfCount("Sheet1", "A1:A10")).toBe(0);
    const after = await adapter.listConditionalFormats("Sheet1", "A1:A10");
    expect(after.ok).toBe(true);
    if (after.ok) expect(after.data).toHaveLength(0);
  });

  it("adds custom CF using rule.formula string (not nested formula.formula)", async () => {
    const adapter = new OfficeJsAdapter();
    const added = await adapter.addConditionalFormat({
      sheetName: "Sheet1",
      range: "B1",
      rule: { kind: "custom", formula: "=A1>5", fillColor: "#00FF00" },
    });
    expect(added.ok).toBe(true);
    if (added.ok) {
      expect(added.data.kind).toBe("custom");
      expect(gates.getCustomFormula("Sheet1", "B1", added.data.id)).toBe("=A1>5");
    }
  });

  it("writes list DV via rule.list, reads back type/rule, clears", async () => {
    const adapter = new OfficeJsAdapter();
    const written = await adapter.writeDataValidation({
      sheetName: "Sheet1",
      range: "C1:C5",
      rule: { type: "list", listValues: ["A", "B", "C"] },
    });
    expect(written.ok).toBe(true);
    if (written.ok) {
      expect(written.data.rule?.type).toBe("list");
      expect(written.data.rule?.listValues).toEqual(["A", "B", "C"]);
    }
    const read = await adapter.readDataValidation("Sheet1", "C1:C5");
    expect(read.ok).toBe(true);
    if (read.ok) expect(read.data.rule?.type).toBe("list");
    const cleared = await adapter.clearDataValidation("Sheet1", "C1:C5");
    expect(cleared.ok).toBe(true);
    const after = await adapter.readDataValidation("Sheet1", "C1:C5");
    expect(after.ok).toBe(true);
    if (after.ok) expect(after.data.rule).toBeNull();
  });

  it("writes wholeNumber DV via rule.wholeNumber with operator", async () => {
    const adapter = new OfficeJsAdapter();
    const written = await adapter.writeDataValidation({
      sheetName: "Sheet1",
      range: "D1",
      rule: {
        type: "wholeNumber",
        operator: "between",
        formula1: "1",
        formula2: "10",
      },
    });
    expect(written.ok).toBe(true);
    if (written.ok) {
      expect(written.data.rule?.type).toBe("wholeNumber");
      expect(written.data.rule?.operator).toBe("between");
      expect(written.data.rule?.formula1).toBe("1");
      expect(written.data.rule?.formula2).toBe("10");
    }
  });
});
