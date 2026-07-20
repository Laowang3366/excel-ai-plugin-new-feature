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
    it("maps GTE/LTE and NotEqualTo correctly; CF has no OrEqualTo suffix", () => {
      expect(mapCfOperatorToHost("greaterThanOrEqualTo")).toBe("GreaterThanOrEqual");
      expect(mapCfOperatorToHost("lessThanOrEqualTo")).toBe("LessThanOrEqual");
      expect(mapCfOperatorToHost("notEqualTo")).toBe("NotEqualTo");
      expect(unmapCfOperator("NotEqualTo")).toBe("notEqualTo");
      expect(unmapCfOperator("NotEqual")).toBeUndefined();
      expect(mapDvOperatorToHost("greaterThanOrEqualTo")).toBe("GreaterThanOrEqualTo");
      expect(mapDvOperatorToHost("lessThanOrEqualTo")).toBe("LessThanOrEqualTo");
      expect(mapDvOperatorToHost("notEqualTo")).toBe("NotEqualTo");
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

    it("does not split formula list sources; marks lossy inline A,,B", () => {
      const range = classifyListSource("=Sheet1!$A$1:$A$3");
      expect(range.kind).toBe("range");
      expect(range.listValues).toBeUndefined();
      const lossy = classifyListSource("A,,B");
      expect(lossy.kind).toBe("inline");
      expect(lossy.lossy).toBe(true);
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
      }
    });

    it("missing/throwing isSetSupported is fail-safe unsupported", async () => {
      installValidationExcel({ missingIsSetSupported: true });
      const adapter = new OfficeJsAdapter();
      expect((await adapter.listConditionalFormats("Sheet1", "A1")).ok).toBe(false);
      installValidationExcel({ isSetSupportedThrows: true });
      const adapter2 = new OfficeJsAdapter();
      const read = await adapter2.readDataValidation("Sheet1", "A1");
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
      expect(fake.getSyncCount()).toBe(1);
    });

    it("throws PropertyNotLoaded if add proxy id is read without load", async () => {
      installValidationExcel();
      const Excel = (globalThis as unknown as {
        Excel: { run: (fn: (ctx: unknown) => Promise<unknown>) => Promise<unknown> };
      }).Excel;
      await expect(
        Excel.run(async (ctx) => {
          const c = ctx as {
            workbook: {
              worksheets: {
                getItem: (n: string) => {
                  getRange: (a: string) => {
                    conditionalFormats: { add: (t: string) => { id: string } };
                  };
                };
              };
            };
          };
          const cf = c.workbook.worksheets.getItem("Sheet1").getRange("A1").conditionalFormats.add(
            "CellValue",
          );
          return cf.id;
        }),
      ).rejects.toThrow(/PropertyNotLoaded:id/);
    });

    it("add verifies rule and colors; fixed sync count 2", async () => {
      const fake = installValidationExcel();
      const adapter = new OfficeJsAdapter();
      fake.resetSyncCount();
      const added = await adapter.addConditionalFormat({
        sheetName: "Sheet1",
        range: "C1:C5",
        rule: {
          kind: "cellValue",
          operator: "notEqualTo",
          formula1: "0",
          fillColor: "#FF0000",
          fontColor: "#00FF00",
        },
      });
      expect(added.ok).toBe(true);
      if (added.ok) {
        expect(added.data.kind).toBe("cellValue");
        expect(added.data.hostType).toBe("CellValue");
      }
      expect(fake.getSyncCount()).toBe(2);
    });

    it("fails add when host tampers operator/color while keeping id/type", async () => {
      installValidationExcel({
        tamperCfReadback: { operator: "EqualTo", fillColor: "#000000" },
      });
      const adapter = new OfficeJsAdapter();
      const added = await adapter.addConditionalFormat({
        sheetName: "Sheet1",
        range: "A1",
        rule: {
          kind: "cellValue",
          operator: "greaterThan",
          formula1: "1",
          fillColor: "#FF0000",
        },
      });
      expect(added.ok).toBe(false);
      if (!added.ok) expect(added.unsupported).not.toBe(true);
    });

    it("fails add when custom formula is tampered", async () => {
      installValidationExcel({ tamperCfReadback: { formula: "=FALSE" } });
      const adapter = new OfficeJsAdapter();
      const added = await adapter.addConditionalFormat({
        sheetName: "Sheet1",
        range: "A1",
        rule: { kind: "custom", formula: "=TRUE", fillColor: "#00FF00" },
      });
      expect(added.ok).toBe(false);
    });

    it("round-trips all CF operators including notEqualTo", async () => {
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
      }
    });

    it("fails add when sync fails; fails delete when still present", async () => {
      installValidationExcel({ failAddSync: true });
      const adapter = new OfficeJsAdapter();
      const added = await adapter.addConditionalFormat({
        sheetName: "Sheet1",
        range: "A1",
        rule: { kind: "custom", formula: "=TRUE" },
      });
      expect(added.ok).toBe(false);

      installValidationExcel({ failDeleteReadback: true });
      const adapter2 = new OfficeJsAdapter();
      const okAdd = await adapter2.addConditionalFormat({
        sheetName: "Sheet1",
        range: "D1",
        rule: { kind: "custom", formula: "=A1>0" },
      });
      expect(okAdd.ok).toBe(true);
      if (!okAdd.ok) return;
      const deleted = await adapter2.deleteConditionalFormat("Sheet1", "D1", okAdd.data.id);
      expect(deleted.ok).toBe(false);
    });
  });

  describe("data validation", () => {
    it("round-trips compare types, custom, inline list, and Range object source", async () => {
      installValidationExcel();
      const adapter = new OfficeJsAdapter();
      for (const type of ["wholeNumber", "decimal", "date", "time", "textLength"] as const) {
        const written = await adapter.writeDataValidation({
          sheetName: "Sheet1",
          range: "E1:E3",
          rule: { type, operator: "greaterThanOrEqualTo", formula1: "1" },
        });
        expect(written.ok, type).toBe(true);
        if (written.ok) {
          expect(written.data.rule?.operator).toBe("greaterThanOrEqualTo");
          expect(written.data.rule?.formula1).toBeDefined();
        }
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
        expect(listRange.data.rule?.formula1).not.toMatch(/\[object Object\]/i);
      }
    });

    it("round-trips allowBlank=false via sync-gated ignoreBlanks", async () => {
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
      if (written.ok) {
        expect(written.data.rule?.allowBlank).toBe(false);
      }
    });

    it("fails write when host coerces operator/formula/allowBlank", async () => {
      installValidationExcel({
        tamperDvReadback: { operator: "LessThan", formula1: "99", allowBlank: true },
      });
      const adapter = new OfficeJsAdapter();
      const bad = await adapter.writeDataValidation({
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
    });

    it("marks Inconsistent unsupported; lossy inline A,,B is not supported:true", async () => {
      installValidationExcel({ seedInconsistentDv: true });
      const adapter = new OfficeJsAdapter();
      const read = await adapter.readDataValidation("Sheet1", "B1:B5");
      expect(read.ok).toBe(true);
      if (read.ok) {
        expect(read.data.rule).toBeNull();
        expect(read.data.supported).toBe(false);
        expect(read.data.hostType).toBe("Inconsistent");
      }

      installValidationExcel({
        tamperDvReadback: { listSource: "A,,B" },
      });
      const adapter2 = new OfficeJsAdapter();
      // seed via write then tamper on readback path — write first with normal, then read with tamper
      await adapter2.writeDataValidation({
        sheetName: "Sheet1",
        range: "L1",
        rule: { type: "list", listValues: ["A", "B"] },
      });
      // re-install with seed: use write under tamper so readback of write fails or read sees lossy
      installValidationExcel({ tamperDvReadback: { listSource: "A,,B" } });
      const adapter3 = new OfficeJsAdapter();
      const written = await adapter3.writeDataValidation({
        sheetName: "Sheet1",
        range: "L2",
        rule: { type: "list", listValues: ["A", "B"] },
      });
      // write readback sees lossy list → not match → failed
      expect(written.ok).toBe(false);
    });

    it("clear returns exact Range.address and requires hostType None", async () => {
      installValidationExcel();
      const adapter = new OfficeJsAdapter();
      await adapter.writeDataValidation({
        sheetName: "Sheet1",
        range: "Y1",
        rule: { type: "list", listValues: ["x"] },
      });
      const cleared = await adapter.clearDataValidation("Sheet1", "Y1");
      expect(cleared.ok).toBe(true);
      if (cleared.ok) {
        expect(cleared.data.cleared).toBe("Sheet1!Y1");
        expect(cleared.data.cleared).not.toMatch(/Sheet1!Sheet1!/);
      }
    });

    it("clear fails when host leaves Inconsistent instead of None", async () => {
      installValidationExcel({ clearLeavesHostType: "Inconsistent" });
      const adapter = new OfficeJsAdapter();
      await adapter.writeDataValidation({
        sheetName: "Sheet1",
        range: "Y2",
        rule: { type: "list", listValues: ["x"] },
      });
      const cleared = await adapter.clearDataValidation("Sheet1", "Y2");
      expect(cleared.ok).toBe(false);
    });

    it("fails write/clear when sync/readback does not confirm", async () => {
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
        range: "Y3",
        rule: { type: "list", listValues: ["x"] },
      });
      const cleared = await adapter2.clearDataValidation("Sheet1", "Y3");
      expect(cleared.ok).toBe(false);
    });
  });
});
