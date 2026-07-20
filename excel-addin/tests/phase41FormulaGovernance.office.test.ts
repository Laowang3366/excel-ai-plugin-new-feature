import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import {
  FORMULA_BACKUP_HEADERS,
  encodeBackupLiteral,
} from "../shared/formulaGovernance";
import { installFormulaGovernanceExcel } from "./fakes/officeJsFormulaGovernanceFake";
import { writeTextMatrix } from "../shared/host/officeJsFormulaGovernanceBackup";
import { withExcel } from "../shared/host/officeJsRuntime";
import { collectFormulaCells } from "../shared/host/officeJsFormulaGovernanceCollect";
import { GOVERNANCE_WRITE_CHUNK } from "../shared/host/officeJsFormulaGovernanceOps";

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

  it("restore fails closed when corrupt data row coexists with valid target backup", async () => {
    const adapter = new OfficeJsAdapter();
    await adapter.convertFormulasToValues({
      scope: "sheet",
      sheetName: "Sheet1",
      backupId: "good-id",
    });
    const b = fake.backupSheet()!;
    // corrupt data row (empty formula) alongside valid rows
    b.values.push(["bad-id", "t", "Sheet1", "Z9", "not-formula", "", "", "", "", ""]);
    const formulasBefore = fake.formulas().map((row) => [...row]);
    const restore = await adapter.restoreFormulas({ backupId: "good-id" });
    expect(restore.ok).toBe(false);
    if (!restore.ok) expect(restore.reason).toMatch(/corrupt|skipped/i);
    // zero write on data sheet
    expect(fake.formulas()).toEqual(formulasBefore);
    // inspect still reports skipped
    const inspect = await adapter.inspectFormulaBackups();
    expect(inspect.ok).toBe(true);
    if (inspect.ok) expect(inspect.data.skippedRows.length).toBeGreaterThanOrEqual(1);
  });

  it("batch spill probe keeps later spillAddress; inspect skips metadata and bounds syncs", async () => {
    // Build many formula cells: first normal, one later is spill parent.
    const formulas: string[][] = [];
    const values: unknown[][] = [];
    for (let r = 0; r < 20; r += 1) {
      formulas.push([`=${r}+1`, r === 10 ? "=SEQUENCE(2)" : `=${r}+2`]);
      values.push([r + 1, r === 10 ? 1 : r + 2]);
    }
    fake = installFormulaGovernanceExcel({
      sheets: [{ name: "Sheet1", formulas, values }],
      spillMap: { "Sheet1!B11": "B11:B12" }, // row 10 0-based -> B11
      excelApi112: true,
    });
    const adapter = new OfficeJsAdapter();

    fake.resetSyncCount();
    const deps = await adapter.inspectFormulaDependencies({
      scope: "sheet",
      sheetName: "Sheet1",
    });
    expect(deps.ok).toBe(true);
    const inspectSyncs = fake.syncCount();
    // Inspect skips locked/spill; must stay near-constant (not ~2*N formula cells).
    expect(inspectSyncs).toBeLessThan(8);

    fake.resetSyncCount();
    const meta = await withExcel("collect-meta", async (context) => {
      const limitations: string[] = [];
      const { cells } = await collectFormulaCells(
        context,
        { scope: "sheet", sheetName: "Sheet1" },
        limitations,
        { includeBackupMetadata: true },
      );
      return { cells, limitations, syncs: fake.syncCount() };
    });
    expect(meta.ok).toBe(true);
    if (meta.ok) {
      // 40 formula cells with per-cell sync would be 80+; batched metadata is one extra sync.
      expect(meta.data.syncs).toBeLessThan(12);
      const spillCell = meta.data.cells.find((c) => c.address === "B11");
      expect((spillCell?.spillAddress ?? "").replace(/\$/g, "")).toMatch(/B11:B12/i);
      // Non-spill first cell must not wipe spill probing for later cells
      const first = meta.data.cells.find((c) => c.address === "A1");
      expect(first?.spillAddress ?? "").toBe("");
    }

    fake.resetSyncCount();
    const conv = await adapter.convertFormulasToValues({
      scope: "sheet",
      sheetName: "Sheet1",
      backupId: "spill1",
    });
    expect(conv.ok).toBe(true);
    // 40 formula cells: write+verify are O(chunks) not O(cells). Upper bound allows collect+backup overhead.
    const convertSyncs = fake.syncCount();
    const formulaCount = 40;
    const maxChunkSyncs = Math.ceil(formulaCount / GOVERNANCE_WRITE_CHUNK) * 2; // write + verify
    expect(convertSyncs).toBeLessThan(maxChunkSyncs + 25);
    // Must stay far below per-cell 2-sync write/verify (would be >= 80 for 40 cells alone).
    expect(convertSyncs).toBeLessThan(formulaCount);

    const b = fake.backupSheet()!;
    const flat = b.values.map((row) => row.map((c) => String(c ?? "")));
    const spillHit = flat.some((row) => row[8] && /B11:B12/i.test(row[8]!));
    expect(spillHit).toBe(true);
  });

  it("repair write/readback is chunk-bounded not per-cell sync", async () => {
    const formulas: string[][] = [];
    const values: unknown[][] = [];
    for (let r = 0; r < 30; r += 1) {
      formulas.push([`=1+#REF!`]);
      values.push(["#REF!"]);
    }
    fake = installFormulaGovernanceExcel({
      sheets: [{ name: "Sheet1", formulas, values }],
    });
    const adapter = new OfficeJsAdapter();
    fake.resetSyncCount();
    const result = await adapter.repairFormulaReferences({
      scope: "sheet",
      sheetName: "Sheet1",
      replacements: [{ find: "#REF!", replace: "Z99" }],
    });
    expect(result.ok).toBe(true);
    const syncs = fake.syncCount();
    const n = 30;
    // write + verify chunks + collect/backup overhead; must not approach 2*n cell syncs.
    expect(syncs).toBeLessThan(n);
    expect(fake.formulas()[0]![0]).toBe("=1+Z99");
    expect(fake.formulas()[29]![0]).toBe("=1+Z99");
  });


  it("excelApi112=false does not probe legacy spill; repair/convert still succeed", async () => {
    fake = installFormulaGovernanceExcel({
      sheets: [
        {
          name: "Sheet1",
          formulas: [
            ["=1+1", "=SEQUENCE(2)"],
            ["", ""],
          ],
          values: [
            [2, 1],
            [null, 2],
          ],
        },
      ],
      spillMap: { "Sheet1!B1": "B1:B2" },
      excelApi112: false,
    });
    const adapter = new OfficeJsAdapter();
    fake.resetSpillProbes();
    const deps = await adapter.inspectFormulaDependencies({
      scope: "sheet",
      sheetName: "Sheet1",
    });
    expect(deps.ok).toBe(true);
    expect(fake.spillLegacyProbeCount()).toBe(0);
    expect(fake.spillNullObjectProbeCount()).toBe(0);

    const conv = await adapter.convertFormulasToValues({
      scope: "sheet",
      sheetName: "Sheet1",
      backupId: "nosill",
    });
    expect(conv.ok).toBe(true);
    if (conv.ok) {
      expect(conv.data.limitations.some((l) => /spillAddress unavailable/i.test(l))).toBe(true);
    }
    expect(fake.spillLegacyProbeCount()).toBe(0);
    // backup spill column empty
    const b = fake.backupSheet()!;
    const spillCol = b.values.slice(2).map((row) => String(row[8] ?? ""));
    expect(spillCol.every((s) => s === "" || s === "undefined")).toBe(true);

    // restore formulas then repair path without legacy spill
    fake.formulas()[0]![0] = "=1+#REF!";
    fake.values()[0]![0] = "#REF!";
    fake.resetSpillProbes();
    const repair = await adapter.repairFormulaReferences({
      scope: "target",
      sheetName: "Sheet1",
      range: "A1",
      replacements: [{ find: "#REF!", replace: "B2" }],
    });
    expect(repair.ok).toBe(true);
    expect(fake.spillLegacyProbeCount()).toBe(0);
    expect(fake.formulas()[0]![0]).toBe("=1+B2");
  });

  it("removeAfterRestore preserves multi-row retained backup multiset", async () => {
    // 3 formula cells -> multi-row backup for keep-me; then one cell for drop-me
    fake = installFormulaGovernanceExcel({
      sheets: [
        {
          name: "Sheet1",
          formulas: [["=1", "=2", "=3"]],
          values: [[1, 2, 3]],
        },
      ],
    });
    const adapter = new OfficeJsAdapter();
    const keep = await adapter.convertFormulasToValues({
      scope: "sheet",
      sheetName: "Sheet1",
      backupId: "keep-multi",
    });
    expect(keep.ok).toBe(true);
    // reseed one formula for second backup id
    fake.formulas()[0]![0] = "=9";
    fake.values()[0]![0] = 9;
    const drop = await adapter.convertFormulasToValues({
      scope: "target",
      sheetName: "Sheet1",
      range: "A1",
      backupId: "drop-one",
    });
    expect(drop.ok).toBe(true);

    const before = fake.backupSheet()!;
    const keepRowsBefore = before.values
      .slice(2)
      .filter((row) => String(row[0] ?? "").includes("keep-multi") || String(row[0] ?? "") === "keep-multi")
      .map((row) => row.map((c) => String(c ?? "")));
    expect(keepRowsBefore.length).toBeGreaterThanOrEqual(3);

    const restored = await adapter.restoreFormulas({
      backupId: "drop-one",
      removeAfterRestore: true,
    });
    expect(restored.ok).toBe(true);

    const after = fake.backupSheet()!;
    const keepRowsAfter = after.values
      .slice(2)
      .filter((row) => {
        const id = String(row[0] ?? "");
        const decoded = id.startsWith("'") ? id.slice(1) : id;
        return decoded === "keep-multi";
      })
      .map((row) => row.map((c) => String(c ?? "")));
    expect(keepRowsAfter.length).toBe(keepRowsBefore.length);
    // field-level: formulas for keep rows still present (encoded or raw)
    for (const row of keepRowsAfter) {
      expect(row[4] || row[4] === "").toBeTruthy();
      const formula = String(row[4] ?? "");
      expect(formula.includes("=") || formula.startsWith("'=")).toBe(true);
    }
    const dropLeft = after.values.slice(2).some((row) => {
      const id = String(row[0] ?? "");
      return id === "drop-one" || id.endsWith("drop-one");
    });
    expect(dropLeft).toBe(false);
  });

});
