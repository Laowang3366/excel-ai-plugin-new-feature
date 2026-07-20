import { afterEach, describe, expect, it } from "vitest";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { oleColorFromHex } from "../shared/host/wpsJsaFormat";
import {
  XL_CELL_VALUE,
  XL_EQUAL,
  XL_GREATER,
  XL_VALIDATE_LIST,
  XL_VALIDATE_WHOLE,
} from "../shared/host/wpsJsaValidationConstants";
import {
  installWpsValidationFake,
  uninstallWpsValidationFake,
} from "./fakes/wpsJsaValidationFake";

describe("phase5 WPS JSA CF/DV", () => {
  afterEach(() => {
    uninstallWpsValidationFake();
  });

  describe("member missing → typed unsupported", () => {
    it("FormatConditions missing", async () => {
      installWpsValidationFake({ missingFormatConditions: true });
      const adapter = new WpsJsaAdapter();
      const list = await adapter.listConditionalFormats("Sheet1", "A1");
      expect(list.ok).toBe(false);
      if (!list.ok) expect(list.unsupported).toBe(true);
    });

    it("Validation missing", async () => {
      installWpsValidationFake({ missingValidation: true });
      const adapter = new WpsJsaAdapter();
      const read = await adapter.readDataValidation("Sheet1", "A1");
      expect(read.ok).toBe(false);
      if (!read.ok) expect(read.unsupported).toBe(true);
    });

    it("FormatConditions.Add missing", async () => {
      installWpsValidationFake({ missingFcAdd: true });
      const adapter = new WpsJsaAdapter();
      const added = await adapter.addConditionalFormat({
        sheetName: "Sheet1",
        range: "A1",
        rule: { kind: "custom", formula: "=TRUE" },
      });
      expect(added.ok).toBe(false);
      if (!added.ok) expect(added.unsupported).toBe(true);
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

  describe("conditional formats", () => {
    it("lists with 1-based index ids; unsupported host types stay unsupported", async () => {
      const fake = installWpsValidationFake();
      fake.seedCondition("Sheet1", "A1", {
        Type: XL_CELL_VALUE,
        Operator: XL_GREATER,
        Formula1: "0",
      });
      fake.seedCondition("Sheet1", "A1", {
        Type: 3, // color scale-like
        Formula1: "x",
      });
      const adapter = new WpsJsaAdapter();
      const listed = await adapter.listConditionalFormats("Sheet1", "A1");
      expect(listed.ok).toBe(true);
      if (!listed.ok) return;
      expect(listed.data).toHaveLength(2);
      expect(listed.data[0]?.id).toBe("1");
      expect(listed.data[0]?.kind).toBe("cellValue");
      expect(listed.data[0]?.supported).toBe(true);
      expect(listed.data[1]?.id).toBe("2");
      expect(listed.data[1]?.kind).toBe("unsupported");
      expect(listed.data[1]?.supported).toBe(false);
    });

    it("adds cellValue/custom with colors and readback; fixed 1-based id", async () => {
      installWpsValidationFake();
      const adapter = new WpsJsaAdapter();
      const added = await adapter.addConditionalFormat({
        sheetName: "Sheet1",
        range: "C1:C5",
        rule: {
          kind: "cellValue",
          operator: "greaterThan",
          formula1: "1",
          fillColor: "#FF0000",
          fontColor: "#00FF00",
        },
      });
      expect(added.ok).toBe(true);
      if (added.ok) {
        expect(added.data.id).toBe("1");
        expect(added.data.hostType).toBe("CellValue");
      }

      const custom = await adapter.addConditionalFormat({
        sheetName: "Sheet1",
        range: "C1:C5",
        rule: { kind: "custom", formula: "=TRUE", fillColor: "#0000FF" },
      });
      expect(custom.ok).toBe(true);
      if (custom.ok) {
        expect(custom.data.id).toBe("2");
        expect(custom.data.kind).toBe("custom");
      }
    });

    it("fails add when host tampers operator/formula/color", async () => {
      installWpsValidationFake({
        tamperCf: { operator: XL_EQUAL, formula1: "99" },
      });
      const bad = await new WpsJsaAdapter().addConditionalFormat({
        sheetName: "Sheet1",
        range: "A1",
        rule: {
          kind: "cellValue",
          operator: "greaterThan",
          formula1: "1",
          fillColor: "#FF0000",
        },
      });
      expect(bad.ok).toBe(false);

      installWpsValidationFake({
        tamperCf: { fillColor: oleColorFromHex("#000000")! },
      });
      const badColor = await new WpsJsaAdapter().addConditionalFormat({
        sheetName: "Sheet1",
        range: "A1",
        rule: {
          kind: "cellValue",
          operator: "equalTo",
          formula1: "1",
          fillColor: "#FF0000",
        },
      });
      expect(badColor.ok).toBe(false);
    });

    it("deletes by 1-based id and verifies remaining multiset", async () => {
      installWpsValidationFake();
      const adapter = new WpsJsaAdapter();
      const a = await adapter.addConditionalFormat({
        sheetName: "Sheet1",
        range: "D1",
        rule: { kind: "custom", formula: "=A1>0" },
      });
      const b = await adapter.addConditionalFormat({
        sheetName: "Sheet1",
        range: "D1",
        rule: { kind: "custom", formula: "=B1>0" },
      });
      const c = await adapter.addConditionalFormat({
        sheetName: "Sheet1",
        range: "D1",
        rule: { kind: "custom", formula: "=C1>0" },
      });
      expect(a.ok && b.ok && c.ok).toBe(true);
      const deleted = await adapter.deleteConditionalFormat("Sheet1", "D1", "2");
      expect(deleted.ok).toBe(true);
      const listed = await adapter.listConditionalFormats("Sheet1", "D1");
      expect(listed.ok).toBe(true);
      if (!listed.ok) return;
      expect(listed.data.map((x) => x.id)).toEqual(["1", "2"]);
      // former 3 is now 2; content preserved by formula fingerprints in delete path
    });

    it("rejects non-between formula2 on add", async () => {
      installWpsValidationFake();
      const bad = await new WpsJsaAdapter().addConditionalFormat({
        sheetName: "Sheet1",
        range: "A1",
        rule: {
          kind: "cellValue",
          operator: "greaterThan",
          formula1: "1",
          formula2: "9",
          fillColor: "#FF0000",
        },
      });
      expect(bad.ok).toBe(false);
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

      // restore path: seed existing, Add throws after Delete
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
      // clearLeavesType makes Delete leave residual — but write also uses Delete.
      // Re-install residual-only scenario after successful seed via direct seed:
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
