import { afterEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { installValidationExcel } from "./fakes/officeJsValidationFake";

describe("phase5 Office.js data validation host paths", () => {
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
    delete (globalThis as { Office?: unknown }).Office;
  });

  it("round-trips all compare types, custom formula, multi/single inline, and Range source", async () => {
    installValidationExcel();
    const adapter = new OfficeJsAdapter();
    for (const type of [
      "wholeNumber",
      "decimal",
      "date",
      "time",
      "textLength",
    ] as const) {
      const written = await adapter.writeDataValidation({
        sheetName: "Sheet1",
        range: "E1:E3",
        rule: { type, operator: "greaterThanOrEqualTo", formula1: "1" },
      });
      expect(written.ok, type).toBe(true);
      if (written.ok) {
        expect(written.data.rule?.type, type).toBe(type);
        expect(written.data.rule?.operator, type).toBe("greaterThanOrEqualTo");
      }
    }

    const custom = await adapter.writeDataValidation({
      sheetName: "Sheet1",
      range: "F1",
      rule: { type: "custom", formula1: "=F1>0" },
    });
    expect(custom.ok).toBe(true);
    if (custom.ok) {
      expect(custom.data.rule?.type).toBe("custom");
      expect(custom.data.rule?.formula1).toMatch(/F1>0/);
    }

    const fake = installValidationExcel();
    const adapter2 = new OfficeJsAdapter();
    fake.resetSyncCount();
    const multi = await adapter2.writeDataValidation({
      sheetName: "Sheet1",
      range: "G9",
      rule: { type: "list", listValues: ["A", "B"] },
    });
    expect(multi.ok).toBe(true);
    if (multi.ok) {
      expect(multi.data.listSourceKind).toBe("inline");
      expect(multi.data.rule?.listValues).toEqual(["A", "B"]);
    }
    expect(fake.getSyncCount()).toBe(2);

    // Anti false-green: single-item list write must succeed with inline readback.
    fake.resetSyncCount();
    const single = await adapter2.writeDataValidation({
      sheetName: "Sheet1",
      range: "G10",
      rule: { type: "list", listValues: ["x"] },
    });
    expect(single.ok).toBe(true);
    if (!single.ok) {
      throw new Error(
        `single listValues write must not fail (got: ${single.reason ?? "failed"})`,
      );
    }
    expect(single.data.listSourceKind).toBe("inline");
    expect(single.data.rule?.listValues).toEqual(["x"]);
    expect(single.data.rule?.listValues).not.toEqual([]);
    expect(fake.getSyncCount()).toBe(2);

    fake.resetSyncCount();
    const ranged = await adapter2.writeDataValidation({
      sheetName: "Sheet1",
      range: "H9",
      rule: { type: "list", formula1: "A1:A3" },
    });
    expect(ranged.ok).toBe(true);
    if (ranged.ok) {
      expect(ranged.data.listSourceKind).toBe("range");
      expect(ranged.data.rule?.formula1).toMatch(/A1:A3/i);
      expect(ranged.data.rule?.listValues).toBeUndefined();
    }
    expect(fake.getSyncCount()).toBe(3);
  });

  it("round-trips allowBlank=false; fails host coercion and custom formula2", async () => {
    installValidationExcel();
    const adapter = new OfficeJsAdapter();
    const written = await adapter.writeDataValidation({
      sheetName: "Sheet1",
      range: "J1",
      rule: {
        type: "wholeNumber",
        operator: "equalTo",
        formula1: "1",
        allowBlank: false,
      },
    });
    expect(written.ok).toBe(true);
    if (written.ok) expect(written.data.rule?.allowBlank).toBe(false);

    installValidationExcel({
      tamperDvReadback: { operator: "LessThan", formula1: "99", allowBlank: true },
    });
    const bad = await new OfficeJsAdapter().writeDataValidation({
      sheetName: "Sheet1",
      range: "K1",
      rule: {
        type: "wholeNumber",
        operator: "greaterThan",
        formula1: "1",
        allowBlank: false,
      },
    });
    expect(bad.ok).toBe(false);

    installValidationExcel({ tamperDvReadback: { formula2: "999" } });
    const badCompare = await new OfficeJsAdapter().writeDataValidation({
      sheetName: "Sheet1",
      range: "M1",
      rule: { type: "wholeNumber", operator: "greaterThan", formula1: "1" },
    });
    expect(badCompare.ok).toBe(false);

    installValidationExcel({ tamperDvReadback: { formula2: "999" } });
    const badCustom = await new OfficeJsAdapter().writeDataValidation({
      sheetName: "Sheet1",
      range: "M2",
      rule: { type: "custom", formula1: "=M2>0" },
    });
    expect(badCustom.ok).toBe(false);
  });

  it("Inconsistent/lossy/=MyList are not supported:true; clear requires None", async () => {
    installValidationExcel({ seedInconsistentDv: true });
    const read = await new OfficeJsAdapter().readDataValidation("Sheet1", "B1:B5");
    expect(read.ok).toBe(true);
    if (read.ok) {
      expect(read.data.rule).toBeNull();
      expect(read.data.supported).toBe(false);
      expect(read.data.hostType).toBe("Inconsistent");
    }

    installValidationExcel({ tamperDvReadback: { listSource: "A,,B" } });
    expect(
      (
        await new OfficeJsAdapter().writeDataValidation({
          sheetName: "Sheet1",
          range: "L2",
          rule: { type: "list", listValues: ["A", "B"] },
        })
      ).ok,
    ).toBe(false);

    installValidationExcel({ tamperDvReadback: { listSource: "=MyList" } });
    const named = await new OfficeJsAdapter().writeDataValidation({
      sheetName: "Sheet1",
      range: "N1",
      rule: { type: "list", listValues: ["A"] },
    });
    expect(named.ok).toBe(false);

    installValidationExcel();
    const adapter = new OfficeJsAdapter();
    const setupClear = await adapter.writeDataValidation({
      sheetName: "Sheet1",
      range: "Y1",
      rule: { type: "list", listValues: ["x"] },
    });
    expect(setupClear.ok).toBe(true);
    const cleared = await adapter.clearDataValidation("Sheet1", "Y1");
    expect(cleared.ok).toBe(true);
    if (cleared.ok) {
      expect(cleared.data.cleared).toBe("Sheet1!Y1");
      expect(cleared.data.cleared).not.toMatch(/Sheet1!Sheet1!/);
    }

    installValidationExcel({ clearLeavesHostType: "Inconsistent" });
    const adapter2 = new OfficeJsAdapter();
    const setupInconsistent = await adapter2.writeDataValidation({
      sheetName: "Sheet1",
      range: "Y2",
      rule: { type: "list", listValues: ["x"] },
    });
    expect(setupInconsistent.ok).toBe(true);
    expect((await adapter2.clearDataValidation("Sheet1", "Y2")).ok).toBe(false);
  });

  it("fails write/clear when sync/readback does not confirm", async () => {
    installValidationExcel({ failWriteSync: true });
    expect(
      (
        await new OfficeJsAdapter().writeDataValidation({
          sheetName: "Sheet1",
          range: "Z1",
          rule: { type: "list", listValues: ["x"] },
        })
      ).ok,
    ).toBe(false);

    installValidationExcel({ failClearReadback: true });
    const adapter2 = new OfficeJsAdapter();
    const setup = await adapter2.writeDataValidation({
      sheetName: "Sheet1",
      range: "Y3",
      rule: { type: "list", listValues: ["x"] },
    });
    expect(setup.ok).toBe(true);
    expect((await adapter2.clearDataValidation("Sheet1", "Y3")).ok).toBe(false);
  });
});
