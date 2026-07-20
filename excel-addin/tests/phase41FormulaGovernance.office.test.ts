import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import {
  FORMULA_BACKUP_HEADERS,
  encodeBackupLiteral,
} from "../shared/formulaGovernance";
import { installFormulaGovernanceExcel } from "./fakes/officeJsFormulaGovernanceFake";
import { writeTextMatrix } from "../shared/host/officeJsFormulaGovernanceBackup";
import { withExcel } from "../shared/host/officeJsRuntime";

describe("phase41 formula governance Office.js", () => {
  let fake: ReturnType<typeof installFormulaGovernanceExcel>;
  beforeEach(() => {
    fake = installFormulaGovernanceExcel({
      sheets: [
        {
          name: "Sheet1",
          formulas: [
            ["=B1", "=1+#REF!"],
            ["='Other'!A1", "x"],
          ],
          values: [
            [1, "#REF!"],
            ["x", "x"],
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

  it("dependencies.inspect keeps text-parse limitations", async () => {
    const result = await new OfficeJsAdapter().inspectFormulaDependencies({
      scope: "sheet",
      sheetName: "Sheet1",
    });
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.data.report.limitations).toContain("text-parse-only");
      expect(result.data.limitations).toContain("no-excel-engine-circularReference");
    }
  });

  it("rejects dimension-mismatched values assignment on backup range", async () => {
    const adapterOk = await withExcel("test", async (context) => {
      const sheet = context.workbook.worksheets.add("_WenggeFormulaBackupDim");
      // 1x1 range cannot accept 1x10
      let threw = false;
      try {
        await writeTextMatrix(context, sheet.getRange("A2"), [[...FORMULA_BACKUP_HEADERS]]);
      } catch (e) {
        threw = /Dimension mismatch/i.test(e instanceof Error ? e.message : String(e));
      }
      return threw;
    });
    expect(adapterOk.ok).toBe(true);
    if (adapterOk.ok) expect(adapterOk.data).toBe(true);
  });

  it("repair incomplete is zero-write; complete backs up formula text as non-evaluated", async () => {
    const adapter = new OfficeJsAdapter();
    const incomplete = await adapter.repairFormulaReferences({
      scope: "sheet",
      sheetName: "Sheet1",
      replacements: [{ find: "NOPE", replace: "A1" }],
    });
    expect(incomplete.ok).toBe(false);
    expect(fake.formulas()[0]![1]).toBe("=1+#REF!");

    const complete = await adapter.repairFormulaReferences({
      scope: "sheet",
      sheetName: "Sheet1",
      replacements: [{ find: "#REF!", replace: "B2" }],
    });
    expect(complete.ok).toBe(true);
    expect(fake.formulas()[0]![1]).toBe("=1+B2");
    const backup = fake.backupSheet();
    expect(backup).toBeTruthy();
    expect(String(backup!.visibility).toLowerCase()).toContain("veryhidden");
    // formula column must not be EVALUATED
    const flat = backup!.values.flat().map(String);
    expect(flat.some((v) => v.includes("EVALUATED"))).toBe(false);
    expect(
      flat.some((v) => v.includes("=1+#REF!") || v.includes(encodeBackupLiteral("=1+#REF!"))),
    ).toBe(true);
  });

  it("convert stores exact formula text including quotes/cross-sheet; restore fail-closed on bad magic", async () => {
    const adapter = new OfficeJsAdapter();
    const conv = await adapter.convertFormulasToValues({
      scope: "target",
      sheetName: "Sheet1",
      range: "A2",
      backupId: "txt1",
    });
    expect(conv.ok).toBe(true);
    const backup = fake.backupSheet()!;
    const asText = backup.values.map((row) => row.map((c) => String(c ?? "")));
    const joined = asText.flat().join("\n");
    expect(joined).toMatch(/Other/);
    expect(joined).not.toContain("EVALUATED");

    // corrupt magic
    backup.values[0]![0] = "NOT_MAGIC";
    backup.formulas[0]![0] = "";
    const badRestore = await adapter.restoreFormulas({ backupId: "txt1" });
    expect(badRestore.ok).toBe(false);
    if (!badRestore.ok) expect(badRestore.reason).toMatch(/corrupt|magic|invalid|not_found/i);
  });

  it("prefix conflict wrong magic is not overwritten; header mismatch restore zero-write", async () => {
    // seed conflicting prefix sheet
    fake = installFormulaGovernanceExcel({
      sheets: [
        {
          name: "Sheet1",
          formulas: [["=1"]],
          values: [[1]],
        },
        {
          name: "_WenggeFormulaBackupConflict",
          formulas: [["x"]],
          values: [["WRONG_MAGIC"]],
        },
      ],
    });
    const adapter = new OfficeJsAdapter();
    const conv = await adapter.convertFormulasToValues({
      scope: "sheet",
      sheetName: "Sheet1",
      backupId: "c1",
    });
    expect(conv.ok).toBe(true);
    // conflict sheet untouched
    const conflict = fake.sheets().find((s) => s.name === "_WenggeFormulaBackupConflict")!;
    expect(String(conflict.values[0]?.[0])).toBe("WRONG_MAGIC");
    // new protocol sheet created (conflict left alone)
    const protocol = fake.sheets().find(
      (s) => s.name.startsWith("_WenggeFormulaBackup") && String(s.values[0]?.[0] ?? "").includes("WENGGE_FORMULA_BACKUP_V1"),
    );
    expect(protocol).toBeTruthy();
    expect(protocol!.name).not.toBe("_WenggeFormulaBackupConflict");

    // header mismatch on real backup
    const b = fake.sheets().find(
      (s) => s.name.startsWith("_WenggeFormulaBackup") && String(s.values[0]?.[0] ?? "").includes("WENGGE_FORMULA_BACKUP_V1"),
    )!;
    // row index 1 is headers
    if (b.values.length < 2) b.values.push([...FORMULA_BACKUP_HEADERS]);
    b.values[1] = ["bad", ...FORMULA_BACKUP_HEADERS.slice(1)];
    const restore = await adapter.restoreFormulas({ backupId: "c1" });
    expect(restore.ok).toBe(false);
    if (!restore.ok) expect(restore.reason).toMatch(/header|corrupt/i);
  });

  it("removeAfterRestore drops only target id and retains others", async () => {
    const adapter = new OfficeJsAdapter();
    await adapter.convertFormulasToValues({
      scope: "sheet",
      sheetName: "Sheet1",
      backupId: "keep-other",
    });
    // manually append second backup id by second convert after re-adding formula
    fake.formulas()[0]![0] = "=99";
    fake.values()[0]![0] = 99;
    await adapter.convertFormulasToValues({
      scope: "target",
      sheetName: "Sheet1",
      range: "A1",
      backupId: "remove-me",
    });
    const restored = await adapter.restoreFormulas({
      backupId: "remove-me",
      removeAfterRestore: true,
    });
    expect(restored.ok).toBe(true);
    const inspect = await adapter.inspectFormulaBackups();
    expect(inspect.ok).toBe(true);
    if (inspect.ok) {
      expect(inspect.data.backups.some((b) => b.backupId === "remove-me")).toBe(false);
      expect(inspect.data.backups.some((b) => b.backupId === "keep-other")).toBe(true);
    }
  });

  it("inspect reports skipped bad rows without failing", async () => {
    const adapter = new OfficeJsAdapter();
    await adapter.convertFormulasToValues({
      scope: "sheet",
      sheetName: "Sheet1",
      backupId: "ok1",
    });
    const b = fake.backupSheet()!;
    // inject empty/corrupt data row
    b.values.push(["", "", "", "", "not-a-formula", "", "", "", "", ""]);
    const inspect = await adapter.inspectFormulaBackups();
    expect(inspect.ok).toBe(true);
    if (inspect.ok) {
      expect(inspect.data.backupCount).toBeGreaterThanOrEqual(1);
      expect(inspect.data.skippedRows.length).toBeGreaterThanOrEqual(1);
    }
  });
});
