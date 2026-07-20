import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { ToolExecutor } from "../shared/tools/executor";
import { TOOL_DEFINITIONS } from "../shared/tools/definitions";
import { dispositionForRisk } from "../shared/agentChat/approvalPolicy";
import { CHAT_READONLY_TOOL_ALLOWLIST } from "../shared/agentChat/chatReadOnlyTools";
import { MockHostAdapter } from "./mockHost";
import { installFormulaGovernanceExcel } from "./fakes/officeJsFormulaGovernanceFake";
import { FORMULA_BACKUP_MAGIC } from "../shared/formulaGovernance";

describe("phase41 formula governance tools", () => {
  describe("registry + policy", () => {
    it("registers five tools with closed schemas and risk levels", () => {
      const names = [
        "formula.dependencies.inspect",
        "formula.references.repair",
        "formula.convertToValues",
        "formula.backups.inspect",
        "formula.backups.restore",
      ] as const;
      for (const name of names) {
        const def = TOOL_DEFINITIONS.find((t) => t.name === name);
        expect(def, name).toBeTruthy();
        expect(def!.parameters.additionalProperties).toBe(false);
      }
      expect(TOOL_DEFINITIONS.find((t) => t.name === "formula.dependencies.inspect")?.riskLevel).toBe(
        "safe",
      );
      expect(TOOL_DEFINITIONS.find((t) => t.name === "formula.backups.inspect")?.riskLevel).toBe(
        "safe",
      );
      expect(TOOL_DEFINITIONS.find((t) => t.name === "formula.references.repair")?.riskLevel).toBe(
        "dangerous",
      );
      expect(dispositionForRisk("dangerous")).toBe("approval");
      expect(CHAT_READONLY_TOOL_ALLOWLIST).toContain("formula.dependencies.inspect");
      expect(CHAT_READONLY_TOOL_ALLOWLIST).toContain("formula.backups.inspect");
      expect(CHAT_READONLY_TOOL_ALLOWLIST).not.toContain("formula.references.repair");
    });

    it("executor rejects unknown fields and incomplete repair plan without write", async () => {
      const host = new MockHostAdapter();
      host.cells.set("Sheet1!A1", {
        values: [["x"]],
        formulas: [["=SUM(#REF!)"]],
      });
      const ex = new ToolExecutor(host);
      const bad = await ex.execute({
        name: "formula.dependencies.inspect",
        arguments: { scope: "sheet", sheetName: "Sheet1", extra: 1 },
      });
      expect(bad.ok).toBe(false);
      if (!bad.ok) expect(bad.error).toMatch(/unknown field/i);

      // Non-matching mapping leaves #REF! → refuse write
      const stillBroken = await ex.execute({
        name: "formula.references.repair",
        arguments: {
          scope: "sheet",
          sheetName: "Sheet1",
          replacements: [{ find: "NOPE", replace: "A1" }],
        },
      });
      expect(stillBroken.ok).toBe(false);
      if (!stillBroken.ok) expect(stillBroken.error).toMatch(/formula_repair_incomplete|incomplete/i);
      expect(host.cells.get("Sheet1!A1")?.formulas[0]?.[0]).toBe("=SUM(#REF!)");
    });
  });

  describe("MockHost executor paths", () => {
    it("inspect dependencies keeps text-parse limitations", async () => {
      const host = new MockHostAdapter();
      host.cells.set("Sheet1!A1", { values: [[1]], formulas: [["=B1"]] });
      host.cells.set("Sheet1!B1", { values: [[2]], formulas: [["=10"]] });
      const ex = new ToolExecutor(host);
      const result = await ex.execute({
        name: "formula.dependencies.inspect",
        arguments: { scope: "sheet", sheetName: "Sheet1" },
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        const data = result.data as {
          report: { limitations: string[]; formulaCount: number };
          limitations: string[];
        };
        expect(data.report.formulaCount).toBe(2);
        expect(data.report.limitations).toContain("text-parse-only");
        expect(data.limitations).toContain("text-parse-only");
        expect(data.limitations).toContain("no-excel-engine-circularReference");
      }
    });

    it("convert backs up then writes values; restore returns formulas", async () => {
      const host = new MockHostAdapter();
      host.cells.set("Sheet1!A1", { values: [[42]], formulas: [["=6*7"]] });
      const ex = new ToolExecutor(host);

      const converted = await ex.execute({
        name: "formula.convertToValues",
        arguments: { scope: "target", sheetName: "Sheet1", range: "A1", backupId: "b1" },
      });
      expect(converted.ok).toBe(true);
      if (converted.ok) {
        const data = converted.data as { backupId: string; convertedFormulaCells: number };
        expect(data.backupId).toBe("b1");
        expect(data.convertedFormulaCells).toBe(1);
      }
      expect(host.cells.get("Sheet1!A1")?.formulas[0]?.[0]).toBe("");
      expect(host.formulaBackupRows.some((r) => r.backupId === "b1" && r.formula === "=6*7")).toBe(
        true,
      );

      const denyNoBackup = await ex.execute({
        name: "formula.convertToValues",
        arguments: {
          scope: "target",
          sheetName: "Sheet1",
          range: "A1",
          createBackup: false,
        },
      });
      expect(denyNoBackup.ok).toBe(false);

      const inspect = await ex.execute({
        name: "formula.backups.inspect",
        arguments: {},
      });
      expect(inspect.ok).toBe(true);
      if (inspect.ok) {
        const data = inspect.data as { backupCount: number; backups: Array<{ backupId: string }> };
        expect(data.backupCount).toBeGreaterThanOrEqual(1);
        expect(data.backups.some((b) => b.backupId === "b1")).toBe(true);
      }

      const restored = await ex.execute({
        name: "formula.backups.restore",
        arguments: { backupId: "b1" },
      });
      expect(restored.ok).toBe(true);
      expect(host.cells.get("Sheet1!A1")?.formulas[0]?.[0]).toBe("=6*7");
      // default retain backup
      expect(host.formulaBackupRows.some((r) => r.backupId === "b1")).toBe(true);

      const missing = await ex.execute({
        name: "formula.backups.restore",
        arguments: { backupId: "no-such" },
      });
      expect(missing.ok).toBe(false);
      if (!missing.ok) expect(missing.error).toMatch(/formula_backup_not_found/);
    });

    it("repair applies mapping only when complete and backs up first", async () => {
      const host = new MockHostAdapter();
      host.cells.set("Sheet1!A1", { values: [[null]], formulas: [["=SUM(#REF!)"]] });
      const ex = new ToolExecutor(host);
      const repaired = await ex.execute({
        name: "formula.references.repair",
        arguments: {
          scope: "sheet",
          sheetName: "Sheet1",
          replacements: [{ find: "#REF!", replace: "B1" }],
        },
      });
      expect(repaired.ok).toBe(true);
      if (repaired.ok) {
        const data = repaired.data as { backupId: string; repairedCount: number; verified: boolean };
        expect(data.repairedCount).toBe(1);
        expect(data.verified).toBe(true);
        expect(data.backupId.length).toBeGreaterThan(0);
      }
      expect(host.cells.get("Sheet1!A1")?.formulas[0]?.[0]).toBe("=SUM(B1)");
      expect(host.formulaBackupRows.some((r) => r.formula === "=SUM(#REF!)")).toBe(true);
    });
  });

  describe("Office.js fake", () => {
    let fake: ReturnType<typeof installFormulaGovernanceExcel>;
    beforeEach(() => {
      fake = installFormulaGovernanceExcel({
        sheets: [
          {
            name: "Sheet1",
            formulas: [
              ["=B1", "=1+#REF!"],
              ["=10", "x"],
            ],
            values: [
              [1, "#REF!"],
              [10, "x"],
            ],
          },
        ],
      });
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      delete (globalThis as { window?: unknown }).window;
    });

    it("dependencies.inspect reports limitations and nodes", async () => {
      const result = await new OfficeJsAdapter().inspectFormulaDependencies({
        scope: "sheet",
        sheetName: "Sheet1",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.report.formulaCount).toBeGreaterThanOrEqual(2);
        expect(result.data.report.limitations).toContain("text-parse-only");
        expect(result.data.limitations).toContain("no-excel-engine-circularReference");
      }
    });

    it("repair incomplete does not write; complete writes with backup", async () => {
      const adapter = new OfficeJsAdapter();
      const incomplete = await adapter.repairFormulaReferences({
        scope: "sheet",
        sheetName: "Sheet1",
        replacements: [{ find: "NOPE", replace: "A1" }],
      });
      expect(incomplete.ok).toBe(false);
      if (!incomplete.ok) expect(incomplete.reason).toMatch(/formula_repair_incomplete|REF/i);
      expect(fake.formulas()[0]![1]).toBe("=1+#REF!");

      const complete = await adapter.repairFormulaReferences({
        scope: "sheet",
        sheetName: "Sheet1",
        replacements: [{ find: "#REF!", replace: "B2" }],
      });
      expect(complete.ok).toBe(true);
      if (complete.ok) {
        expect(complete.data.verified).toBe(true);
        expect(complete.data.backupId.length).toBeGreaterThan(0);
      }
      expect(fake.formulas()[0]![1]).toBe("=1+B2");
      const backup = fake.backupSheet();
      expect(backup).toBeTruthy();
      expect(backup!.formulas[0]?.[0] || backup!.values[0]?.[0]).toBeTruthy();
      // magic in values A1 after write (values path)
      // our fake stores backup via values assignment
      const a1 = String(backup!.values[0]?.[0] ?? backup!.formulas[0]?.[0] ?? "");
      // encode uses values write for backup rows; magic row may be in values
      expect(a1 === FORMULA_BACKUP_MAGIC || backup!.name.startsWith("_WenggeFormulaBackup")).toBe(
        true,
      );
    });

    it("convertToValues backs up then clears formulas", async () => {
      const adapter = new OfficeJsAdapter();
      const result = await adapter.convertFormulasToValues({
        scope: "target",
        sheetName: "Sheet1",
        range: "A1",
        backupId: "conv1",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.backupId).toBe("conv1");
        expect(result.data.convertedFormulaCells).toBe(1);
        expect(result.data.verified).toBe(true);
      }
      expect(fake.formulas()[0]![0]).toBe("");
      expect(fake.backupSheet()).toBeTruthy();
    });

    it("backups.inspect and restore round-trip", async () => {
      const adapter = new OfficeJsAdapter();
      await adapter.convertFormulasToValues({
        scope: "sheet",
        sheetName: "Sheet1",
        backupId: "rt1",
      });
      const inspected = await adapter.inspectFormulaBackups();
      expect(inspected.ok).toBe(true);
      if (inspected.ok) {
        expect(inspected.data.backupCount).toBeGreaterThanOrEqual(1);
        expect(inspected.data.backups.some((b) => b.backupId === "rt1")).toBe(true);
      }
      const restored = await adapter.restoreFormulas({ backupId: "rt1" });
      expect(restored.ok).toBe(true);
      if (restored.ok) {
        expect(restored.data.restoredCount).toBeGreaterThanOrEqual(1);
        expect(restored.data.verified).toBe(true);
      }
    });
  });

  describe("WPS member-check paths", () => {
    afterEach(() => {
      delete (globalThis as { window?: unknown }).window;
      delete (globalThis as { Application?: unknown }).Application;
    });

    function installWpsApp(opts?: { withAdd?: boolean; withUsedRange?: boolean }) {
      const withAdd = opts?.withAdd !== false;
      const withUsedRange = opts?.withUsedRange !== false;
      const cells = new Map<string, { formula: string; value: unknown; numberFormat: string }>();
      cells.set("A1", { formula: "=B1", value: 1, numberFormat: "General" });
      cells.set("B1", { formula: "=1+#REF!", value: "#REF!", numberFormat: "General" });

      const sheets: any[] = [];
      const makeSheet = (name: string) => {
        const sheet: any = {
          Name: name,
          Range(address: string) {
            const bare = address.replace(/\$/g, "").toUpperCase().split(":")[0]!;
            // multi-cell for used range A1:B1
            if (address.includes(":")) {
              return {
                Address: `${name}!${address}`,
                get Formula() {
                  return [
                    [cells.get("A1")?.formula ?? "", cells.get("B1")?.formula ?? ""],
                  ];
                },
                set Formula(v: unknown) {
                  // ignore block write in this simple fake
                  void v;
                },
                get Value2() {
                  return [[cells.get("A1")?.value ?? null, cells.get("B1")?.value ?? null]];
                },
                set Value2(v: unknown) {
                  void v;
                },
              };
            }
            const cell = cells.get(bare) ?? {
              formula: "",
              value: null,
              numberFormat: "General",
            };
            cells.set(bare, cell);
            return {
              Address: `${name}!${bare}`,
              get Formula() {
                return cell.formula;
              },
              set Formula(v: unknown) {
                cell.formula = typeof v === "string" ? v : String(v ?? "");
              },
              get Value2() {
                return cell.value;
              },
              set Value2(v: unknown) {
                cell.value = v;
                cell.formula = "";
              },
              get NumberFormat() {
                return cell.numberFormat;
              },
              set NumberFormat(v: unknown) {
                cell.numberFormat = String(v ?? "");
              },
              Locked: false,
            };
          },
          get UsedRange() {
            if (!withUsedRange) return undefined;
            return {
              Address: "A1:B1",
              get Value2() {
                return [
                  [cells.get("A1")?.value ?? null, cells.get("B1")?.value ?? null],
                ];
              },
              get Formula() {
                return [
                  [cells.get("A1")?.formula ?? "", cells.get("B1")?.formula ?? ""],
                ];
              },
            };
          },
        };
        return sheet;
      };
      sheets.push(makeSheet("Sheet1"));
      const workbook = {
        Name: "Book1",
        ActiveSheet: sheets[0],
        Worksheets: {
          Count: sheets.length,
          Item(indexOrName: number | string) {
            if (typeof indexOrName === "number") return sheets[indexOrName - 1];
            const hit = sheets.find((s) => s.Name === indexOrName);
            if (!hit) throw new Error("sheet missing");
            return hit;
          },
          Add: withAdd
            ? () => {
                const s = makeSheet(`_WenggeFormulaBackup${sheets.length}`);
                sheets.push(s);
                workbook.Worksheets.Count = sheets.length;
                return s;
              }
            : undefined,
        },
      };
      const app = { ActiveWorkbook: workbook, Name: "WPS" };
      (globalThis as any).window = { Application: app };
      (globalThis as any).Application = app;
      return { cells, sheets, workbook };
    }

    it("inspects dependencies when Range/UsedRange available", async () => {
      installWpsApp();
      const result = await new WpsJsaAdapter().inspectFormulaDependencies({
        scope: "sheet",
        sheetName: "Sheet1",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.report.limitations).toContain("text-parse-only");
        expect(result.data.report.formulaCount).toBeGreaterThanOrEqual(1);
      }
    });

    it("returns typed unsupported when Worksheets.Add missing for convert backup", async () => {
      installWpsApp({ withAdd: false });
      const result = await new WpsJsaAdapter().convertFormulasToValues({
        scope: "sheet",
        sheetName: "Sheet1",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        // either fail with Add unavailable or unsupported
        expect(
          result.unsupported === true || /Add|backup/i.test(result.reason),
        ).toBe(true);
      }
    });

    it("repairs with backup when Add available", async () => {
      const env = installWpsApp();
      const result = await new WpsJsaAdapter().repairFormulaReferences({
        scope: "sheet",
        sheetName: "Sheet1",
        replacements: [{ find: "#REF!", replace: "A1" }],
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.verified).toBe(true);
        expect(result.data.backupId.length).toBeGreaterThan(0);
      }
      expect(env.cells.get("B1")?.formula).toBe("=1+A1");
    });
  });
});
