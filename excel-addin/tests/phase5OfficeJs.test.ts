import { afterEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { installValidationExcel } from "./fakes/officeJsValidationFake";

describe("phase5 Office.js CF/DV host paths", () => {
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
    delete (globalThis as { Office?: unknown }).Office;
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
      if (!read.ok) expect(read.unsupported).toBe(true);
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
          const cf = c.workbook.worksheets
            .getItem("Sheet1")
            .getRange("A1")
            .conditionalFormats.add("CellValue");
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
      if (added.ok) expect(added.data.hostType).toBe("CellValue");
      expect(fake.getSyncCount()).toBe(2);
    });

    it("fails add when host tampers operator/color/formula while keeping id/type", async () => {
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

      installValidationExcel({ tamperCfReadback: { formula: "=FALSE" } });
      const adapter2 = new OfficeJsAdapter();
      const custom = await adapter2.addConditionalFormat({
        sheetName: "Sheet1",
        range: "A1",
        rule: { kind: "custom", formula: "=TRUE", fillColor: "#00FF00" },
      });
      expect(custom.ok).toBe(false);
    });

    it("round-trips CF operators; fails on non-between host formula2", async () => {
      installValidationExcel();
      const adapter = new OfficeJsAdapter();
      for (const operator of [
        "greaterThan",
        "greaterThanOrEqualTo",
        "equalTo",
        "notEqualTo",
        "between",
      ] as const) {
        const added = await adapter.addConditionalFormat({
          sheetName: "Sheet1",
          range: "C1:C5",
          rule: {
            kind: "cellValue",
            operator,
            formula1: "1",
            formula2: operator === "between" ? "9" : undefined,
            fillColor: "#FF0000",
          },
        });
        expect(added.ok, operator).toBe(true);
      }

      installValidationExcel({ tamperCfReadback: { formula2: "999" } });
      const bad = await new OfficeJsAdapter().addConditionalFormat({
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
    it("round-trips types and Range source; Range needs extra fixed sync", async () => {
      installValidationExcel();
      const adapter = new OfficeJsAdapter();
      for (const type of ["wholeNumber", "decimal", "textLength"] as const) {
        const written = await adapter.writeDataValidation({
          sheetName: "Sheet1",
          range: "E1:E3",
          rule: { type, operator: "greaterThanOrEqualTo", formula1: "1" },
        });
        expect(written.ok, type).toBe(true);
      }

      const fake = installValidationExcel();
      const adapter2 = new OfficeJsAdapter();
      fake.resetSyncCount();
      const inline = await adapter2.writeDataValidation({
        sheetName: "Sheet1",
        range: "G9",
        rule: { type: "list", listValues: ["A", "B"] },
      });
      expect(inline.ok).toBe(true);
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

      installValidationExcel({ clearLeavesHostType: "Inconsistent" });
      const adapter2 = new OfficeJsAdapter();
      await adapter2.writeDataValidation({
        sheetName: "Sheet1",
        range: "Y2",
        rule: { type: "list", listValues: ["x"] },
      });
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
      await adapter2.writeDataValidation({
        sheetName: "Sheet1",
        range: "Y3",
        rule: { type: "list", listValues: ["x"] },
      });
      expect((await adapter2.clearDataValidation("Sheet1", "Y3")).ok).toBe(false);
    });
  });
});
