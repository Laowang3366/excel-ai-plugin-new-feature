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

    it("round-trips all 8 CF operators; fails on non-between host formula2", async () => {
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
            formula2:
              operator === "between" || operator === "notBetween" ? "9" : undefined,
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
});
