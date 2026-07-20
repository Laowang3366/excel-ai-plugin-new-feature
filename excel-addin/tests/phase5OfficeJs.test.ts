import { afterEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import {
  mapCfOperatorToHost,
  mapDvOperatorToHost,
  unmapCfOperator,
  unmapDvOperator,
  classifyCfHostType,
  classifyListSource,
} from "../shared/host/officeJsValidationMapping";
import { installValidationExcel } from "./fakes/officeJsValidationFake";

describe("phase5 Office.js CF/DV hardened", () => {
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
    delete (globalThis as { Office?: unknown }).Office;
  });

  describe("operator maps (exact, separate CF vs DV)", () => {
    it("maps GTE/LTE correctly and does not confuse CF with DV host tokens", () => {
      expect(mapCfOperatorToHost("greaterThanOrEqualTo")).toBe("GreaterThanOrEqual");
      expect(mapCfOperatorToHost("lessThanOrEqualTo")).toBe("LessThanOrEqual");
      expect(mapDvOperatorToHost("greaterThanOrEqualTo")).toBe("GreaterThanOrEqualTo");
      expect(mapDvOperatorToHost("lessThanOrEqualTo")).toBe("LessThanOrEqualTo");
      expect(unmapCfOperator("GreaterThanOrEqual")).toBe("greaterThanOrEqualTo");
      expect(unmapCfOperator("GreaterThan")).toBe("greaterThan");
      expect(unmapDvOperator("GreaterThanOrEqualTo")).toBe("greaterThanOrEqualTo");
      // CF host token must not parse as DV's OrEqualTo form via accidental includes
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
      expect(classifyCfHostType("Custom").kind).toBe("custom");
      expect(classifyCfHostType("DataBar").hostType).toBe("DataBar");
    });

    it("does not split formula list sources into listValues", () => {
      const range = classifyListSource("=Sheet1!$A$1:$A$3");
      expect(range.kind).toBe("range");
      expect(range.listValues).toBeUndefined();
      expect(range.formula1).toContain("Sheet1");
      const inline = classifyListSource('"A,B,C"');
      expect(inline.kind).toBe("inline");
      expect(inline.listValues).toEqual(["A", "B", "C"]);
    });
  });

  describe("requirement-set precheck", () => {
    it("CF tools are typed unsupported when ExcelApi 1.6 is false", async () => {
      installValidationExcel({ excelApi16: false, excelApi18: true });
      const adapter = new OfficeJsAdapter();
      const list = await adapter.listConditionalFormats("Sheet1", "A1");
      expect(list.ok).toBe(false);
      if (!list.ok) {
        expect(list.unsupported).toBe(true);
        expect(list.reason).toMatch(/ExcelApi 1\.6/);
      }
    });

    it("DV tools are typed unsupported when ExcelApi 1.8 is false", async () => {
      installValidationExcel({ excelApi16: true, excelApi18: false });
      const adapter = new OfficeJsAdapter();
      const read = await adapter.readDataValidation("Sheet1", "A1");
      expect(read.ok).toBe(false);
      if (!read.ok) {
        expect(read.unsupported).toBe(true);
        expect(read.reason).toMatch(/ExcelApi 1\.8/);
      }
    });

    it("missing isSetSupported is fail-safe unsupported", async () => {
      installValidationExcel({ missingIsSetSupported: true });
      const adapter = new OfficeJsAdapter();
      const list = await adapter.listConditionalFormats("Sheet1", "A1");
      expect(list.ok).toBe(false);
      if (!list.ok) expect(list.unsupported).toBe(true);
    });

    it("isSetSupported throw is fail-safe unsupported", async () => {
      installValidationExcel({ isSetSupportedThrows: true });
      const adapter = new OfficeJsAdapter();
      const read = await adapter.readDataValidation("Sheet1", "A1");
      expect(read.ok).toBe(false);
      if (!read.ok) expect(read.unsupported).toBe(true);
    });
  });

  describe("conditional formats", () => {
    it("lists many rules with O(1) sync (not per-rule)", async () => {
      const fake = installValidationExcel({ seedManyCf: 40 });
      const adapter = new OfficeJsAdapter();
      fake.resetSyncCount();
      const listed = await adapter.listConditionalFormats("Sheet1", "A1:A100");
      expect(listed.ok).toBe(true);
      if (!listed.ok) return;
      expect(listed.data.length).toBe(40);
      expect(listed.data.some((r) => r.hostType === "DataBar" && r.kind === "unsupported")).toBe(
        true,
      );
      expect(listed.data.some((r) => r.kind === "cellValue" && r.supported)).toBe(true);
      // one Excel.run → one sync for list batch
      expect(fake.getSyncCount()).toBe(1);
    });

    it("does not mislabel ContainsText as cellValue", async () => {
      installValidationExcel({ seedContainsText: true });
      const adapter = new OfficeJsAdapter();
      const listed = await adapter.listConditionalFormats("Sheet1", "A1:A10");
      expect(listed.ok).toBe(true);
      if (!listed.ok) return;
      const contains = listed.data.find((r) => r.id === "cf_contains");
      expect(contains?.kind).toBe("unsupported");
      expect(contains?.hostType).toBe("ContainsText");
      expect(contains?.supported).toBe(false);
    });

    it("round-trips full CF operators via cellValue.rule whole assign", async () => {
      installValidationExcel();
      const adapter = new OfficeJsAdapter();
      for (const operator of [
        "greaterThan",
        "greaterThanOrEqualTo",
        "lessThan",
        "lessThanOrEqualTo",
        "equalTo",
        "notEqualTo",
        "between",
        "notBetween",
      ] as const) {
        const added = await adapter.addConditionalFormat({
          sheetName: "Sheet1",
          range: "C1:C5",
          rule: {
            kind: "cellValue",
            operator,
            formula1: "1",
            formula2: operator === "between" || operator === "notBetween" ? "9" : undefined,
            fillColor: "#FF0000",
          },
        });
        expect(added.ok, operator).toBe(true);
        if (added.ok) {
          expect(added.data.kind).toBe("cellValue");
          expect(added.data.hostType).toBe("CellValue");
          expect(added.data.supported).toBe(true);
        }
      }
    });

    it("fails add when sync fails (not success without host commit)", async () => {
      installValidationExcel({ failAddSync: true });
      const adapter = new OfficeJsAdapter();
      const added = await adapter.addConditionalFormat({
        sheetName: "Sheet1",
        range: "A1",
        rule: { kind: "custom", formula: "=TRUE", fillColor: "#00FF00" },
      });
      expect(added.ok).toBe(false);
      if (!added.ok) expect(added.unsupported).not.toBe(true);
    });

    it("fails delete when host still has the rule after sync", async () => {
      installValidationExcel({ failDeleteReadback: true });
      const adapter = new OfficeJsAdapter();
      const added = await adapter.addConditionalFormat({
        sheetName: "Sheet1",
        range: "D1",
        rule: { kind: "custom", formula: "=A1>0" },
      });
      expect(added.ok).toBe(true);
      if (!added.ok) return;
      const deleted = await adapter.deleteConditionalFormat("Sheet1", "D1", added.data.id);
      expect(deleted.ok).toBe(false);
    });
  });

  describe("data validation", () => {
    it("round-trips compare types and keeps formula list source as range kind", async () => {
      installValidationExcel();
      const adapter = new OfficeJsAdapter();
      for (const type of ["wholeNumber", "decimal", "date", "time", "textLength"] as const) {
        const written = await adapter.writeDataValidation({
          sheetName: "Sheet1",
          range: "E1:E3",
          rule: {
            type,
            operator: "greaterThanOrEqualTo",
            formula1: "1",
          },
        });
        expect(written.ok, type).toBe(true);
        if (written.ok) {
          expect(written.data.rule?.type).toBe(type);
          expect(written.data.rule?.operator).toBe("greaterThanOrEqualTo");
          expect(written.data.supported).toBe(true);
        }
      }

      const custom = await adapter.writeDataValidation({
        sheetName: "Sheet1",
        range: "F1",
        rule: { type: "custom", formula1: "=F1>0" },
      });
      expect(custom.ok).toBe(true);

      const listInline = await adapter.writeDataValidation({
        sheetName: "Sheet1",
        range: "G1",
        rule: { type: "list", listValues: ["A", "B"] },
      });
      expect(listInline.ok).toBe(true);
      if (listInline.ok) {
        expect(listInline.data.listSourceKind).toBe("inline");
        expect(listInline.data.rule?.listValues).toEqual(["A", "B"]);
      }

      const listRange = await adapter.writeDataValidation({
        sheetName: "Sheet1",
        range: "H1",
        rule: { type: "list", formula1: "Sheet1!A1:A3" },
      });
      expect(listRange.ok).toBe(true);
      if (listRange.ok) {
        expect(listRange.data.listSourceKind).toBe("range");
        expect(listRange.data.rule?.listValues).toBeUndefined();
        expect(listRange.data.rule?.formula1).toMatch(/A1:A3/i);
      }
    });

    it("marks Inconsistent as unsupported state without a writable rule", async () => {
      installValidationExcel({ seedInconsistentDv: true });
      const adapter = new OfficeJsAdapter();
      const read = await adapter.readDataValidation("Sheet1", "B1:B5");
      expect(read.ok).toBe(true);
      if (!read.ok) return;
      expect(read.data.rule).toBeNull();
      expect(read.data.supported).toBe(false);
      expect(read.data.hostType).toBe("Inconsistent");
      expect(read.data.limitations?.join(" ")).toMatch(/Inconsistent/i);
    });

    it("fails write/clear when readback does not confirm", async () => {
      installValidationExcel({ failWriteSync: true });
      const adapter = new OfficeJsAdapter();
      const bad = await adapter.writeDataValidation({
        sheetName: "Sheet1",
        range: "Z1",
        rule: { type: "list", listValues: ["x"] },
      });
      expect(bad.ok).toBe(false);

      installValidationExcel({ failClearReadback: true });
      const adapter2 = new OfficeJsAdapter();
      await adapter2.writeDataValidation({
        sheetName: "Sheet1",
        range: "Y1",
        rule: { type: "list", listValues: ["x"] },
      });
      const cleared = await adapter2.clearDataValidation("Sheet1", "Y1");
      // clear sync succeeds but readback still sees rule → fail
      expect(cleared.ok).toBe(false);
    });
  });
});
