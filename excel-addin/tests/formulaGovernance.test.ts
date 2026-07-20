import { describe, expect, it } from "vitest";
import { strictDecodeBackup } from "../shared/host/officeJsFormulaGovernanceBackup";
import {
  DEPENDENCY_LIMITATIONS,
  FORMULA_BACKUP_HEADERS,
  FORMULA_BACKUP_MAGIC,
  buildDependencyReport,
  classifyFormula,
  createBackupRows,
  decodeBackupSheet,
  decodeBackupLiteral,
  encodeBackupSheet,
  encodeBackupLiteral,
  findCycles,
  isDynamicArray,
  isFormula,
  leadingFunction,
  makeCellId,
  normalizeA1Address,
  parseFormulaReferences,
  planFormulaRepairs,
  planRestore,
  removeStringLiterals,
  summarizeBackups,
  tryClassifyFormula,
  validateRepairedFormulas,
  type FormulaCellRecord,
  type FormulaEdge,
} from "../shared/formulaGovernance";

describe("address + literals", () => {
  it("normalizes absolute refs and builds cell ids", () => {
    expect(normalizeA1Address("$B$2")).toBe("B2");
    expect(normalizeA1Address("Sheet1!$a$1:$b$2")).toBe("A1:B2");
    expect(makeCellId("Sheet1", "$A$1")).toBe("Sheet1!A1");
  });

  it("strips string literals including escaped quotes", () => {
    expect(removeStringLiterals('=A1&"hello""world"&B1')).toBe('=A1&""&B1');
  });
});

describe("parseFormulaReferences", () => {
  it("parses local absolute and relative A1", () => {
    const refs = parseFormulaReferences("=$A$1+B2", "Data");
    expect(refs.map((r) => r.targetId).sort()).toEqual([
      "Data!A1",
      "Data!B2",
    ]);
    expect(refs.every((r) => r.kind === "same-sheet")).toBe(true);
  });

  it("parses cross-sheet including spaced sheet names", () => {
    const refs = parseFormulaReferences("='My Sheet'!$C$3+Other!A1", "Here");
    const ids = refs.map((r) => r.targetId).sort();
    expect(ids).toContain("My Sheet!C3");
    expect(ids).toContain("Other!A1");
    expect(refs.filter((r) => r.kind === "cross-sheet")).toHaveLength(2);
  });

  it("parses external workbook refs", () => {
    const refs = parseFormulaReferences("=[Budget.xlsx]Summary!B2", "Here");
    expect(refs).toHaveLength(1);
    expect(refs[0]!.kind).toBe("external");
    expect(refs[0]!.targetId).toBe("external:[Budget.xlsx]Summary!B2");
  });

  it("does not treat quoted A1 inside strings as refs", () => {
    const refs = parseFormulaReferences('="A1"&"Sheet1!B2"', "Here");
    expect(refs).toHaveLength(0);
  });

  it("dedupes repeated references", () => {
    const refs = parseFormulaReferences("=A1+A1+$A$1", "S");
    expect(refs.filter((r) => r.targetId === "S!A1")).toHaveLength(1);
  });
});

describe("dependency graph", () => {
  const cells = (rows: FormulaCellRecord[]) => rows;

  it("builds nodes/edges/dependents and always includes limitations", () => {
    const report = buildDependencyReport(
      cells([
        { sheetName: "S", address: "A1", formula: "=B1" },
        { sheetName: "S", address: "B1", formula: "=10" },
      ]),
    );
    expect(report.formulaCount).toBe(2);
    expect(report.limitations).toEqual([...DEPENDENCY_LIMITATIONS]);
    expect(report.limitations).toContain("text-parse-only");
    expect(report.limitations).toContain("no-excel-engine-circularReference");
    const a1 = report.nodes.find((n) => n.id === "S!A1")!;
    expect(a1.precedents).toContain("S!B1");
    const b1 = report.nodes.find((n) => n.id === "S!B1")!;
    expect(b1.dependents).toContain("S!A1");
  });

  it("detects #REF! brokenReferences", () => {
    const report = buildDependencyReport(
      cells([{ sheetName: "S", address: "A1", formula: "=SUM(#REF!)" }]),
    );
    expect(report.brokenReferences).toEqual([
      { cell: "S!A1", formula: "=SUM(#REF!)", reason: "#REF!" },
    ]);
  });

  it("detects self-reference cycle", () => {
    const report = buildDependencyReport(
      cells([{ sheetName: "S", address: "A1", formula: "=A1" }]),
    );
    expect(report.cycles.length).toBeGreaterThan(0);
    expect(report.cycles[0]!.path.join("->").toUpperCase()).toContain("S!A1");
  });

  it("detects multi-node cycles and dedupes edges", () => {
    const report = buildDependencyReport(
      cells([
        { sheetName: "S", address: "A1", formula: "=B1" },
        { sheetName: "S", address: "B1", formula: "=A1+A1" },
      ]),
    );
    expect(report.edgeCount).toBe(2); // A→B, B→A (deduped A1 twice)
    expect(report.cycles.length).toBeGreaterThan(0);
    const path = report.cycles[0]!.path.map((p) => p.toUpperCase());
    expect(path).toContain("S!A1");
    expect(path).toContain("S!B1");
  });

  it("findCycles helper ignores edges to non-nodes", () => {
    const edges: FormulaEdge[] = [
      { from: "S!A1", to: "S!Z9", kind: "same-sheet", reference: "Z9" },
    ];
    expect(findCycles(["S!A1"], edges)).toEqual([]);
  });
});

describe("repair plan", () => {
  it("applies explicit replacements and marks complete", () => {
    const plan = planFormulaRepairs(
      [{ sheetName: "S", address: "A1", formula: "=SUM(#REF!)" }],
      [{ find: "#REF!", replace: "B:B" }],
    );
    expect(plan.complete).toBe(true);
    expect(plan.repairedCount).toBe(1);
    expect(plan.repairs[0]!.after).toBe("=SUM(B:B)");
    expect(plan.unresolvedCount).toBe(0);
  });

  it("leaves residual #REF! as incomplete / unresolved", () => {
    const plan = planFormulaRepairs(
      [{ sheetName: "S", address: "A1", formula: "=#REF!+#REF!" }],
      [{ find: "#REF!", replace: "#REF!" }],
    );
    // after still contains #REF!
    expect(plan.complete).toBe(false);
    expect(plan.unresolvedCount).toBeGreaterThan(0);
  });

  it("skips healthy formulas unless applyAllMappings", () => {
    const cells = [
      { sheetName: "S", address: "A1", formula: "=A2" },
      { sheetName: "S", address: "B1", formula: "=#REF!" },
    ];
    const limited = planFormulaRepairs(cells, [{ find: "A2", replace: "C2" }]);
    expect(limited.repairedCount).toBe(0);
    expect(limited.unresolvedCount).toBe(1); // B1 unchanged still broken

    const all = planFormulaRepairs(
      cells,
      [{ find: "A2", replace: "C2" }],
      { applyAllMappings: true },
    );
    expect(all.repairs.some((r) => r.after === "=C2")).toBe(true);
  });

  it("validateRepairedFormulas reports residual breaks", () => {
    expect(
      validateRepairedFormulas([{ cell: "S!A1", formula: "=A1" }]),
    ).toEqual({ ok: true });
    expect(
      validateRepairedFormulas([{ cell: "S!A1", formula: "=#REF!" }]),
    ).toEqual({ ok: false, stillBroken: ["S!A1"] });
  });
});

describe("backup V1 codec", () => {
  it("round-trips rows through encode/decode", () => {
    const rows = createBackupRows(
      [
        {
          sheetName: "Data",
          address: "$A$1",
          formula: "=SUM(B1:B10)",
          formulaR1C1: "=SUM(R1C2:R10C2)",
          numberFormat: "0.00",
          locked: true,
          spillAddress: "A1:A5",
        },
      ],
      { backupId: "bk1", createdAt: "2026-01-01T00:00:00.000Z", sourceRange: "Data!A1" },
    );
    const grid = encodeBackupSheet(rows);
    expect(grid[0]![0]).toBe(FORMULA_BACKUP_MAGIC);
    expect(grid[1]).toEqual([...FORMULA_BACKUP_HEADERS]);
    const decoded = decodeBackupSheet(grid);
    expect(decoded.ok).toBe(true);
    expect(decoded.grid!.rows).toHaveLength(1);
    expect(decoded.grid!.rows[0]).toMatchObject({
      backupId: "bk1",
      sheet: "Data",
      address: "A1",
      formula: "=SUM(B1:B10)",
      locked: true,
      spillAddress: "A1:A5",
    });
  });

  it("groups multiple backupIds and plans restore", () => {
    const rows = [
      ...createBackupRows(
        [{ sheetName: "S", address: "A1", formula: "=1" }],
        { backupId: "a", createdAt: "2026-01-02T00:00:00.000Z" },
      ),
      ...createBackupRows(
        [
          { sheetName: "S", address: "B1", formula: "=2" },
          { sheetName: "T", address: "A1", formula: "=3" },
        ],
        { backupId: "b", createdAt: "2026-01-03T00:00:00.000Z" },
      ),
    ];
    const summaries = summarizeBackups(rows);
    expect(summaries.map((s) => s.backupId).sort()).toEqual(["a", "b"]);
    const b = summaries.find((s) => s.backupId === "b")!;
    expect(b.formulaCount).toBe(2);
    expect(b.sheets.sort()).toEqual(["S", "T"]);

    const plan = planRestore(rows, "b");
    expect("error" in plan).toBe(false);
    if ("error" in plan) return;
    expect(plan.items).toHaveLength(2);
    expect(planRestore(rows, "missing")).toEqual({
      error: "formula_backup_not_found",
    });
  });

  it("skips corrupt data rows and rejects bad magic", () => {
    const badMagic = decodeBackupSheet([["NOPE"], [...FORMULA_BACKUP_HEADERS]]);
    expect(badMagic.ok).toBe(false);

    const mixed = decodeBackupSheet([
      [FORMULA_BACKUP_MAGIC],
      [...FORMULA_BACKUP_HEADERS],
      ["id1", "t", "S", "A1", "=1", "", "", "0", "", ""],
      ["", "t", "S", "A2", "=2", "", "", "0", "", ""], // missing id
      ["id2", "t", "", "B1", "=3", "", "", "0", "", ""], // missing sheet
      ["id3", "t", "S", "C1", "not-formula", "", "", "0", "", ""],
    ]);
    expect(mixed.grid!.rows).toHaveLength(1);
    expect(mixed.skipped.length).toBeGreaterThan(0);
  });
});

describe("classification", () => {
  it("classifies plain / dynamic / legacyArray", () => {
    expect(classifyFormula("=SUM(A1:A10)")).toBe("plain");
    expect(classifyFormula("=FILTER(A:A,A:A<>\"\")")).toBe("dynamic");
    expect(classifyFormula("=A1:A3*2")).toBe("dynamic");
    expect(classifyFormula("=SUM(A1:A10)", true)).toBe("legacyArray");
    expect(tryClassifyFormula("90")).toBeNull();
    expect(isFormula("=1")).toBe(true);
    expect(leadingFunction("=XLOOKUP(A1,B:B,C:C)")).toBe("XLOOKUP");
    expect(isDynamicArray("=TRANSPOSE(A1:A3)")).toBe(true);
  });

  it("does not treat quoted FILTER text as modern function", () => {
    expect(classifyFormula('="FILTER(A:A)"')).toBe("plain");
  });

  it("throws on non-formula classify", () => {
    expect(() => classifyFormula("nope")).toThrow(/公式/);
  });
});

describe("backup formula text literals", () => {
  it("encodes leading = so sheet values store text, and decodes round-trip", () => {
    const formula = "='Other Sheet'!A1&\"x\"";
    const enc = encodeBackupLiteral(formula);
    expect(enc.startsWith("'=")).toBe(true);
    expect(decodeBackupLiteral(enc)).toBe(formula);
    expect(decodeBackupLiteral(formula)).toBe(formula); // host may strip apostrophe
    const rows = createBackupRows(
      [{ sheetName: "S", address: "A1", formula, formulaR1C1: "=R[0]C[1]" }],
      { backupId: "b1", sourceRange: "A1" },
    );
    const grid = encodeBackupSheet(rows);
    expect(grid[2]![4]).toBe(enc);
    const decoded = decodeBackupSheet(grid);
    expect(decoded.ok).toBe(true);
    expect(decoded.grid?.rows[0]?.formula).toBe(formula);
    expect(decoded.grid?.rows[0]?.formulaR1C1).toBe("=R[0]C[1]");
  });
});

describe("strictDecodeBackup restore fail-closed", () => {
  it("rejects any skipped data row even when magic/header and other rows are valid", () => {
    const matrix = [
      [FORMULA_BACKUP_MAGIC],
      [...FORMULA_BACKUP_HEADERS],
      [
        "id1",
        "t",
        "Sheet1",
        "A1",
        "'=A2",
        "",
        "General",
        "false",
        "",
        "sheet",
      ],
      ["bad", "t", "Sheet1", "Z9", "not-a-formula", "", "", "", "", ""],
    ];
    const strict = strictDecodeBackup(matrix);
    expect(strict.ok).toBe(false);
    if (!strict.ok) {
      expect(strict.error).toMatch(/corrupt|skipped/i);
      expect(strict.skipped.length).toBeGreaterThanOrEqual(1);
    }
    // inspect-oriented decode still returns partial rows
    const decoded = decodeBackupSheet(matrix);
    expect(decoded.grid?.rows.some((r) => r.backupId === "id1")).toBe(true);
    expect(decoded.skipped.length).toBeGreaterThanOrEqual(1);
  });
});
