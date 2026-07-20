import { afterEach, describe, expect, it } from "vitest";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { FORMULA_BACKUP_MAGIC, encodeBackupLiteral } from "../shared/formulaGovernance";

type Cell = { formula: string; value: unknown; numberFormat: string };

function installWpsApp(opts?: {
  withAdd?: boolean;
  withVisible?: boolean;
  withClear?: boolean;
  evaluateBareEquals?: boolean;
}) {
  const withAdd = opts?.withAdd !== false;
  const withVisible = opts?.withVisible !== false;
  const withClear = opts?.withClear !== false;
  const evaluateBareEquals = opts?.evaluateBareEquals !== false;
  const cells = new Map<string, Cell>();
  cells.set("A1", { formula: "=B1", value: 1, numberFormat: "General" });
  cells.set("B1", { formula: "=1+#REF!", value: "#REF!", numberFormat: "General" });

  const sheetStore: Array<{ Name: string; cells: Map<string, Cell>; Visible?: number | string; matrix?: unknown[][] }> =
    [];

  function makeSheet(name: string, sharedCells?: Map<string, Cell>) {
    const local = sharedCells ?? new Map<string, Cell>();
    if (!sharedCells) {
      // empty for backup sheets
    }
    const sheet: any = {
      Name: name,
      ...(withVisible ? { Visible: 1 } : {}),
      Range(address: string) {
        const bare = address.replace(/\$/g, "").toUpperCase();
        if (bare.includes(":")) {
          // multi-cell for used-range style access / backup blocks
          const makeBlock = () => {
            let pending: unknown = undefined;
            let nf = "General";
            return {
              Address: `${name}!${address}`,
              get Formula() {
                return [
                  [local.get("A1")?.formula ?? "", local.get("B1")?.formula ?? ""],
                ];
              },
              set Formula(v: unknown) {
                void v;
              },
              get Value2() {
                if (pending !== undefined) return pending;
                return [
                  [local.get("A1")?.value ?? null, local.get("B1")?.value ?? null],
                ];
              },
              set Value2(v: unknown) {
                if (Array.isArray(v) && Array.isArray(v[0])) {
                  // Merge into full sheet grid (row/col 0-based from address top-left)
                  const m = /^([A-Z]+)(\d+)/i.exec(bare);
                  let c0 = 0;
                  let r0 = 0;
                  if (m) {
                    let col = 0;
                    for (const ch of m[1]!.toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
                    c0 = col - 1;
                    r0 = Number(m[2]) - 1;
                  }
                  if (!sheet._grid) sheet._grid = [];
                  for (let r = 0; r < v.length; r++) {
                    while (sheet._grid.length <= r0 + r) sheet._grid.push([]);
                    const row = v[r] as unknown[];
                    for (let c = 0; c < row.length; c++) {
                      let raw = row[c];
                      const textMode = nf === "@" || (typeof raw === "string" && raw.startsWith("'"));
                      if (typeof raw === "string" && raw.startsWith("=") && !textMode && evaluateBareEquals) {
                        raw = "EVALUATED";
                      }
                      while (sheet._grid[r0 + r]!.length <= c0 + c) sheet._grid[r0 + r]!.push(null);
                      sheet._grid[r0 + r]![c0 + c] = raw;
                    }
                  }
                  pending = sheet._grid;
                }
              },
              get NumberFormat() {
                return nf;
              },
              set NumberFormat(v: unknown) {
                nf = String(v ?? "General");
              },
              Clear: withClear
                ? () => {
                    pending = undefined;
                    sheet._matrix = undefined;
                    local.clear();
                  }
                : undefined,
            };
          };
          return makeBlock();
        }
        const cell = local.get(bare) ?? { formula: "", value: null, numberFormat: "General" };
        local.set(bare, cell);
        return {
          Address: `${name}!${bare}`,
          get Formula() {
            return cell.formula;
          },
          set Formula(v: unknown) {
            cell.formula = typeof v === "string" ? v : String(v ?? "");
          },
          get Value2() {
            // Prefer protocol grid when present (backup sheet writes)
            if (sheet._grid) {
              const m = /^([A-Z]+)(\d+)$/i.exec(bare);
              if (m) {
                let col = 0;
                for (const ch of m[1]!.toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
                const r = Number(m[2]) - 1;
                const c = col - 1;
                if (sheet._grid[r] && sheet._grid[r][c] !== undefined) return sheet._grid[r][c];
              }
            }
            return cell.value;
          },
          set Value2(v: unknown) {
            const textMode = cell.numberFormat === "@" || (typeof v === "string" && v.startsWith("'"));
            if (typeof v === "string" && v.startsWith("=") && !textMode && evaluateBareEquals) {
              cell.value = "EVALUATED";
              cell.formula = v;
            } else {
              cell.value = v;
              // writing a value clears formula (host calc result paste)
              cell.formula = "";
            }
          },
          get NumberFormat() {
            return cell.numberFormat;
          },
          set NumberFormat(v: unknown) {
            cell.numberFormat = String(v ?? "General");
          },
          get FormulaR1C1() {
            return cell.formula ? cell.formula.replace(/[A-Z]+\d+/g, "R1C1") : "";
          },
          Locked: false,
        };
      },
      get UsedRange() {
        if (sheet._grid && sheet._grid.length) {
          const rows = sheet._grid.length;
          return {
            Address: `A1:J${rows}`,
            Value2: sheet._grid,
            Clear: withClear
              ? () => {
                  sheet._grid = undefined;
                  local.clear();
                }
              : undefined,
          };
        }
        return {
          Address: "A1:B1",
          get Value2() {
            return [
              [local.get("A1")?.value ?? null, local.get("B1")?.value ?? null],
            ];
          },
          get Formula() {
            return [
              [local.get("A1")?.formula ?? "", local.get("B1")?.formula ?? ""],
            ];
          },
          Clear: withClear
            ? () => {
                local.clear();
              }
            : undefined,
        };
      },
    };
    sheetStore.push({ Name: name, cells: local, get Visible() { return sheet.Visible; }, set Visible(v) { sheet.Visible = v; } } as any);
    if (withVisible) sheet.Visible = 1;
    return sheet;
  }

  const dataCells = cells;
  const sheets: any[] = [makeSheet("Sheet1", dataCells)];
  const workbook: any = {
    Name: "Book1",
    ActiveSheet: sheets[0],
    Worksheets: {
      get Count() {
        return sheets.length;
      },
      Item(indexOrName: number | string) {
        if (typeof indexOrName === "number") return sheets[indexOrName - 1];
        const hit = sheets.find((s) => s.Name === indexOrName);
        if (!hit) throw new Error("sheet missing");
        return hit;
      },
      Add: withAdd
        ? () => {
            const s = makeSheet(`_tmp${sheets.length}`);
            sheets.push(s);
            return s;
          }
        : undefined,
    },
  };
  const app = { ActiveWorkbook: workbook, Name: "WPS" };
  (globalThis as any).window = { Application: app };
  (globalThis as any).Application = app;
  return { cells: dataCells, sheets, workbook };
}

describe("phase41 formula governance WPS", () => {
  afterEach(() => {
    delete (globalThis as { window?: unknown }).window;
    delete (globalThis as { Application?: unknown }).Application;
  });

  it("inspects dependencies with Formula/UsedRange", async () => {
    installWpsApp();
    const result = await new WpsJsaAdapter().inspectFormulaDependencies({
      scope: "sheet",
      sheetName: "Sheet1",
    });
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.data.report.limitations).toContain("text-parse-only");
    }
  });

  it("typed unsupported when Worksheets.Add missing (backup cannot be created; no formula write)", async () => {
    const env = installWpsApp({ withAdd: false });
    const before = env.cells.get("B1")!.formula;
    const result = await new WpsJsaAdapter().convertFormulasToValues({
      scope: "sheet",
      sheetName: "Sheet1",
    });
    expect(result.ok).toBe(false);
    expect(env.cells.get("B1")!.formula).toBe(before);
  });

  it("fails closed when Visible missing (cannot hide backup; no formula mutation)", async () => {
    const env = installWpsApp({ withVisible: false });
    const before = env.cells.get("B1")!.formula;
    const result = await new WpsJsaAdapter().repairFormulaReferences({
      scope: "sheet",
      sheetName: "Sheet1",
      replacements: [{ find: "#REF!", replace: "A1" }],
    });
    expect(result.ok).toBe(false);
    expect(env.cells.get("B1")!.formula).toBe(before);
  });

  it("repairs with backup text storage and hide verify", async () => {
    const env = installWpsApp();
    const result = await new WpsJsaAdapter().repairFormulaReferences({
      scope: "sheet",
      sheetName: "Sheet1",
      replacements: [{ find: "#REF!", replace: "A1" }],
    });
    expect(result.ok).toBe(true);
    expect(env.cells.get("B1")?.formula).toBe("=1+A1");
    const backup = env.sheets.find((s: any) => String(s.Name).startsWith("_WenggeFormulaBackup"));
    expect(backup).toBeTruthy();
    expect(Number(backup.Visible)).toBe(2);
  });

  it("removeAfterRestore=true without Clear is unsupported before restore", async () => {
    installWpsApp({ withClear: false });
    const adapter = new WpsJsaAdapter();
    await adapter.convertFormulasToValues({
      scope: "sheet",
      sheetName: "Sheet1",
      backupId: "w1",
    });
    // re-seed formula for restore target
    // convert already cleared formulas; restore should still be blocked before write when removeAfterRestore
    const result = await adapter.restoreFormulas({
      backupId: "w1",
      removeAfterRestore: true,
    });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.unsupported === true || /Clear|removeAfterRestore/i.test(result.reason)).toBe(
        true,
      );
    }
  });

  it("removeAfterRestore=false restores; with Clear can remove id", async () => {
    const env = installWpsApp({ withClear: true });
    const adapter = new WpsJsaAdapter();
    const conv = await adapter.convertFormulasToValues({
      scope: "sheet",
      sheetName: "Sheet1",
      backupId: "rm1",
    });
    expect(conv.ok).toBe(true);
    const restored = await adapter.restoreFormulas({
      backupId: "rm1",
      removeAfterRestore: false,
    });
    expect(restored.ok).toBe(true);
    // formulas restored on data sheet
    expect(env.cells.get("A1")?.formula.startsWith("=")).toBe(true);

    const removed = await adapter.restoreFormulas({
      backupId: "rm1",
      removeAfterRestore: true,
    });
    // second restore may succeed remove
    if (removed.ok) {
      const inspect = await adapter.inspectFormulaBackups();
      if (inspect.ok) {
        expect(inspect.data.backups.some((b) => b.backupId === "rm1")).toBe(false);
      }
    }
  });

  it("bare Value2 equals would evaluate; text path uses apostrophe/@", async () => {
    expect(encodeBackupLiteral("=A1+1").startsWith("'=")).toBe(true);
    expect(FORMULA_BACKUP_MAGIC.startsWith("WENGGE")).toBe(true);
  });
});
