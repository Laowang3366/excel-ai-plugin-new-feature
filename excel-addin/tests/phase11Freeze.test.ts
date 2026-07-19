import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { ToolExecutor, TOOL_DEFINITIONS } from "../shared/tools";
import { installFreezeExcel } from "./fakes/officeJsFreezeFake";
import { MockHostAdapter } from "./mockHost";

describe("phase11 sheet.freeze", () => {
  it("registers freeze tools", () => {
    const names = TOOL_DEFINITIONS.map((tool) => tool.name);
    expect(names).toContain("sheet.freeze.get");
    expect(names).toContain("sheet.freeze.set");
  });

  describe("Office.js sync-gated writeback", () => {
    let gates: ReturnType<typeof installFreezeExcel>;

    beforeEach(() => {
      gates = installFreezeExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
    });

    it("gets none and writebacks rows/columns/at/clear after sync", async () => {
      const adapter = new OfficeJsAdapter();
      const none = await adapter.getSheetFreeze("Sheet1");
      expect(none.ok).toBe(true);
      if (none.ok) {
        expect(none.data).toEqual({
          sheetName: "Sheet1",
          address: null,
          rowCount: 0,
          columnCount: 0,
        });
      }

      const rows = await adapter.setSheetFreeze({
        sheetName: "Sheet1",
        command: "rows",
        count: 2,
      });
      expect(rows.ok).toBe(true);
      if (rows.ok) {
        expect(rows.data.rowCount).toBe(2);
        expect(rows.data.columnCount).toBe(0);
        expect(rows.data.address).toContain("A3");
      }
      expect(gates.getCommitted("Sheet1")?.rowCount).toBe(2);

      const cols = await adapter.setSheetFreeze({
        sheetName: "Sheet2",
        command: "columns",
        count: 1,
      });
      expect(cols.ok).toBe(true);
      if (cols.ok) {
        expect(cols.data.columnCount).toBe(1);
        expect(cols.data.address).toContain("A1");
      }

      const at = await adapter.setSheetFreeze({
        sheetName: "Sheet1",
        command: "at",
        address: "C5",
      });
      expect(at.ok).toBe(true);
      if (at.ok) expect(at.data.address).toContain("C5");

      const cleared = await adapter.setSheetFreeze({
        sheetName: "Sheet1",
        command: "clear",
      });
      expect(cleared.ok).toBe(true);
      if (cleared.ok) {
        expect(cleared.data.address).toBeNull();
        expect(cleared.data.rowCount).toBe(0);
        expect(cleared.data.columnCount).toBe(0);
      }
    });

    it("location stays stale until sync (proves writeback depends on sync)", async () => {
      const sheet = (
        globalThis as unknown as {
          Excel: { run: <T>(fn: (ctx: { workbook: { worksheets: { getItem: (n: string) => unknown } }; sync: () => Promise<void> }) => Promise<T>) => Promise<T> };
        }
      ).Excel;

      await sheet.run(async (context) => {
        const ws = context.workbook.worksheets.getItem("Sheet1") as {
          freezePanes: {
            freezeRows: (n: number) => void;
            getLocationOrNullObject: () => {
              isNullObject: boolean;
              rowCount: number;
              load: (p: string) => void;
            };
          };
        };
        ws.freezePanes.freezeRows(3);
        const before = ws.freezePanes.getLocationOrNullObject();
        before.load("rowCount");
        // Intentionally no sync yet: committed location must still be empty.
        expect(before.isNullObject).toBe(true);
        expect(before.rowCount).toBe(0);
        expect(gates.getPending("Sheet1")?.rowCount).toBe(3);
        expect(gates.getCommitted("Sheet1")).toBeNull();

        await context.sync();
        const after = ws.freezePanes.getLocationOrNullObject();
        after.load("rowCount");
        expect(after.isNullObject).toBe(false);
        expect(after.rowCount).toBe(3);
        expect(gates.getCommitted("Sheet1")?.rowCount).toBe(3);
      });
    });
  });

  it("executor success and validation branches", async () => {
    const executor = new ToolExecutor(new MockHostAdapter());

    expect(
      (await executor.execute({ name: "sheet.freeze.get", arguments: { sheetName: "Sheet1" } }))
        .ok,
    ).toBe(true);

    expect(
      (
        await executor.execute({
          name: "sheet.freeze.set",
          arguments: { sheetName: "Sheet1", command: "rows", count: 3 },
        })
      ).ok,
    ).toBe(true);

    for (const args of [
      { sheetName: "Sheet1", command: "ghost" },
      { sheetName: "Sheet1", command: "rows" },
      { sheetName: "Sheet1", command: "rows", count: 0 },
      { sheetName: "Sheet1", command: "at" },
      { sheetName: "Sheet1", command: "at", address: "" },
      { sheetName: "Sheet1", command: "clear", count: 1 },
      { sheetName: "Sheet1", command: "rows", count: null },
      { sheetName: null, command: "clear" },
    ]) {
      const result = await executor.execute({
        name: "sheet.freeze.set",
        arguments: args as Record<string, unknown>,
      });
      expect(result.ok).toBe(false);
    }
  });

  it("WPS returns unsupported for freeze get/set", async () => {
    const executor = new ToolExecutor(new WpsJsaAdapter());
    for (const name of ["sheet.freeze.get", "sheet.freeze.set"] as const) {
      const result = await executor.execute({
        name,
        arguments:
          name === "sheet.freeze.get"
            ? { sheetName: "Sheet1" }
            : { sheetName: "Sheet1", command: "clear" },
      });
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.unsupported).toBe(true);
    }
  });
});
