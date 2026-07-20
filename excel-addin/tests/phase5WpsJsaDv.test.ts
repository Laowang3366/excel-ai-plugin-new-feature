import { afterEach, describe, expect, it } from "vitest";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import {
  XL_BETWEEN,
  XL_EQUAL,
  XL_VALIDATE_LIST,
  XL_VALIDATE_WHOLE,
} from "../shared/host/wpsJsaValidationConstants";
import {
  installWpsValidationFake,
  uninstallWpsValidationFake,
} from "./fakes/wpsJsaValidationFake";

describe("phase5 WPS JSA data validation", () => {
  afterEach(() => {
    uninstallWpsValidationFake();
  });

  describe("member missing → typed unsupported", () => {
    it("Validation missing", async () => {
      installWpsValidationFake({ missingValidation: true });
      const adapter = new WpsJsaAdapter();
      const read = await adapter.readDataValidation("Sheet1", "A1");
      expect(read.ok).toBe(false);
      if (!read.ok) expect(read.unsupported).toBe(true);
    });

    it("Validation.Add missing", async () => {
      installWpsValidationFake({ missingValidationAdd: true });
      const adapter = new WpsJsaAdapter();
      const written = await adapter.writeDataValidation({
        sheetName: "Sheet1",
        range: "A1",
        rule: { type: "list", listValues: ["x"] },
      });
      expect(written.ok).toBe(false);
      if (!written.ok) expect(written.unsupported).toBe(true);
    });
  });

  describe("data validation", () => {
    it("round-trips compare types, custom, single/multi inline, range source", async () => {
      installWpsValidationFake();
      const adapter = new WpsJsaAdapter();
      for (const type of [
        "wholeNumber",
        "decimal",
        "date",
        "time",
        "textLength",
      ] as const) {
        const written = await adapter.writeDataValidation({
          sheetName: "Sheet1",
          range: "E1",
          rule: { type, operator: "greaterThanOrEqualTo", formula1: "1" },
        });
        expect(written.ok, type).toBe(true);
      }

      const custom = await adapter.writeDataValidation({
        sheetName: "Sheet1",
        range: "F1",
        rule: { type: "custom", formula1: "=F1>0" },
      });
      expect(custom.ok).toBe(true);

      const multi = await adapter.writeDataValidation({
        sheetName: "Sheet1",
        range: "G1",
        rule: { type: "list", listValues: ["A", "B"] },
      });
      expect(multi.ok).toBe(true);
      if (multi.ok) {
        expect(multi.data.listSourceKind).toBe("inline");
        expect(multi.data.rule?.listValues).toEqual(["A", "B"]);
      }

      const single = await adapter.writeDataValidation({
        sheetName: "Sheet1",
        range: "G2",
        rule: { type: "list", listValues: ["x"] },
      });
      expect(single.ok).toBe(true);
      if (!single.ok) throw new Error("single list write must succeed");
      expect(single.data.listSourceKind).toBe("inline");
      expect(single.data.rule?.listValues).toEqual(["x"]);

      const ranged = await adapter.writeDataValidation({
        sheetName: "Sheet1",
        range: "H1",
        rule: { type: "list", formula1: "A1:A3" },
      });
      expect(ranged.ok).toBe(true);
      if (ranged.ok) {
        expect(ranged.data.listSourceKind).toBe("range");
        expect(ranged.data.rule?.listValues).toBeUndefined();
      }
    });

    it("fails when host tampers rule; restores previous on Add failure", async () => {
      installWpsValidationFake({
        tamperDv: { operator: XL_EQUAL, formula1: "99" },
      });
      const bad = await new WpsJsaAdapter().writeDataValidation({
        sheetName: "Sheet1",
        range: "K1",
        rule: {
          type: "wholeNumber",
          operator: "greaterThan",
          formula1: "1",
        },
      });
      expect(bad.ok).toBe(false);

      const fake = installWpsValidationFake({ addThrows: "add boom" });
      fake.seedValidation("Sheet1", "R1", {
        Type: XL_VALIDATE_LIST,
        Formula1: "Old",
        IgnoreBlank: true,
        InCellDropdown: true,
      });
      const adapter = new WpsJsaAdapter();
      const failed = await adapter.writeDataValidation({
        sheetName: "Sheet1",
        range: "R1",
        rule: { type: "list", listValues: ["New"] },
      });
      expect(failed.ok).toBe(false);
      const restored = fake.getValidation("Sheet1", "R1");
      expect(restored?.Type).toBe(XL_VALIDATE_LIST);
      expect(restored?.Formula1).toBe("Old");
    });

    it("clear requires None; residual host type fails; setup must succeed", async () => {
      installWpsValidationFake();
      const adapter = new WpsJsaAdapter();
      const setup = await adapter.writeDataValidation({
        sheetName: "Sheet1",
        range: "Y1",
        rule: { type: "list", listValues: ["x"] },
      });
      expect(setup.ok).toBe(true);
      const cleared = await adapter.clearDataValidation("Sheet1", "Y1");
      expect(cleared.ok).toBe(true);
      if (cleared.ok) expect(cleared.data.cleared).toMatch(/Y1/);

      installWpsValidationFake({ clearLeavesType: XL_VALIDATE_WHOLE });
      const adapter2 = new WpsJsaAdapter();
      const setup2 = await adapter2.writeDataValidation({
        sheetName: "Sheet1",
        range: "Y2",
        rule: { type: "list", listValues: ["x"] },
      });
      expect(setup2.ok).toBe(true);

      const fake = installWpsValidationFake({ clearLeavesType: XL_VALIDATE_WHOLE });
      fake.seedValidation("Sheet1", "Y3", {
        Type: XL_VALIDATE_LIST,
        Formula1: "x",
        IgnoreBlank: true,
        InCellDropdown: true,
      });
      const badClear = await new WpsJsaAdapter().clearDataValidation("Sheet1", "Y3");
      expect(badClear.ok).toBe(false);
    });

    it("between without Formula2 is unsupported read; between round-trip works", async () => {
      const fake = installWpsValidationFake();
      fake.seedValidation("Sheet1", "B1", {
        Type: XL_VALIDATE_WHOLE,
        Operator: XL_BETWEEN,
        Formula1: "1",
        IgnoreBlank: true,
        InCellDropdown: false,
      });
      const adapter = new WpsJsaAdapter();
      const read = await adapter.readDataValidation("Sheet1", "B1");
      expect(read.ok).toBe(true);
      if (!read.ok) return;
      expect(read.data.rule).toBeNull();
      expect(read.data.supported).toBe(false);
      expect(read.data.limitations?.join(" ") ?? "").toMatch(/missing formula2/i);

      installWpsValidationFake();
      const adapter2 = new WpsJsaAdapter();
      const written = await adapter2.writeDataValidation({
        sheetName: "Sheet1",
        range: "B2",
        rule: {
          type: "wholeNumber",
          operator: "between",
          formula1: "1",
          formula2: "10",
        },
      });
      expect(written.ok).toBe(true);
      if (written.ok) {
        expect(written.data.supported).toBe(true);
        expect(written.data.rule?.operator).toBe("between");
        expect(written.data.rule?.formula1).toBe("1");
        expect(written.data.rule?.formula2).toBe("10");
      }
    });

    it("Delete failure with existing rule aborts before Add and preserves original", async () => {
      const fake = installWpsValidationFake({ validationDeleteThrows: "delete boom" });
      fake.seedValidation("Sheet1", "D1", {
        Type: XL_VALIDATE_LIST,
        Formula1: "KeepMe",
        IgnoreBlank: true,
        InCellDropdown: true,
      });
      const adapter = new WpsJsaAdapter();
      const failed = await adapter.writeDataValidation({
        sheetName: "Sheet1",
        range: "D1",
        rule: { type: "list", listValues: ["New"] },
      });
      expect(failed.ok).toBe(false);
      if (!failed.ok) {
        expect(failed.unsupported).toBeFalsy();
        expect(failed.reason).toMatch(/Delete failed|left original/i);
      }
      const kept = fake.getValidation("Sheet1", "D1");
      expect(kept?.Type).toBe(XL_VALIDATE_LIST);
      expect(kept?.Formula1).toBe("KeepMe");
    });

    it("rejects non-between formula2; rejects =MyList as range kind", async () => {
      installWpsValidationFake();
      const adapter = new WpsJsaAdapter();
      const extra = await adapter.writeDataValidation({
        sheetName: "Sheet1",
        range: "M1",
        rule: {
          type: "wholeNumber",
          operator: "greaterThan",
          formula1: "1",
          formula2: "9",
        },
      });
      expect(extra.ok).toBe(false);

      const named = await adapter.writeDataValidation({
        sheetName: "Sheet1",
        range: "N1",
        rule: { type: "list", formula1: "=MyList" },
      });
      expect(named.ok).toBe(false);
    });
  });
});
