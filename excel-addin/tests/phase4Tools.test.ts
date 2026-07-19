import { describe, expect, it } from "vitest";
import { ToolExecutor, TOOL_DEFINITIONS } from "../shared/tools";
import { MockHostAdapter } from "./mockHost";

describe("phase4 tools", () => {
  it("registers expand, formula.context, sheet.operation", () => {
    const names = TOOL_DEFINITIONS.map((tool) => tool.name);
    expect(names).toContain("formula.context");
    expect(names).toContain("sheet.operation");
    const read = TOOL_DEFINITIONS.find((tool) => tool.name === "range.read");
    expect(JSON.stringify(read?.parameters)).toContain("spill");
  });

  it("reads expand modes and formula context", async () => {
    const host = new MockHostAdapter();
    await host.writeFormulas("Sheet1", "A1", [["=SUM(1,2)"]]);
    const executor = new ToolExecutor(host);
    for (const expand of ["spill", "currentArray", "currentRegion"] as const) {
      const result = await executor.execute({
        name: "range.read",
        arguments: { sheetName: "Sheet1", range: "A1", expand },
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data).toMatchObject({ expanded: true, expandMode: expand });
      }
    }
    const ctx = await executor.execute({
      name: "formula.context",
      arguments: { sheetName: "Sheet1", range: "A1" },
    });
    expect(ctx.ok).toBe(true);
    if (ctx.ok) {
      const data = ctx.data as {
        address: string;
        formulas: { address: string; formula: string }[];
      };
      expect(data.address).toBeTruthy();
      expect(data.formulas.length).toBe(1);
      expect(data.formulas[0]?.address).toBe("A1");
      expect(data.formulas[0]?.formula).toBe("=SUM(1,2)");
    }

    const offset = await executor.execute({
      name: "formula.context",
      arguments: { sheetName: "Sheet1", range: "D5" },
    });
    // Mock stores per key; write formula at D5 first
    await host.writeFormulas("Sheet1", "D5", [["=9"]]);
    const offset2 = await executor.execute({
      name: "formula.context",
      arguments: { sheetName: "Sheet1", range: "D5" },
    });
    expect(offset2.ok).toBe(true);
    if (offset2.ok) {
      const data = offset2.data as { formulas: { address: string }[] };
      expect(data.formulas[0]?.address).toBe("D5");
    }
    void offset;
  });

  it("runs sheet.operation add/rename/delete/copy/move with 1-based position", async () => {
    const host = new MockHostAdapter();
    const executor = new ToolExecutor(host);
    expect(
      (
        await executor.execute({
          name: "sheet.operation",
          arguments: { operation: "add", sheetName: "Data" },
        })
      ).ok,
    ).toBe(true);
    expect(
      (
        await executor.execute({
          name: "sheet.operation",
          arguments: { operation: "rename", sheetName: "Data", newName: "Metrics" },
        })
      ).ok,
    ).toBe(true);
    expect(
      (
        await executor.execute({
          name: "sheet.operation",
          arguments: { operation: "copy", sheetName: "Metrics", newName: "Metrics2" },
        })
      ).ok,
    ).toBe(true);
    expect(
      (
        await executor.execute({
          name: "sheet.operation",
          arguments: { operation: "move", sheetName: "Metrics2", position: 1 },
        })
      ).ok,
    ).toBe(true);
    expect(
      (
        await executor.execute({
          name: "sheet.operation",
          arguments: { operation: "delete", sheetName: "Metrics2" },
        })
      ).ok,
    ).toBe(true);
  });

  it("rejects invalid expand/operation/position", async () => {
    const host = new MockHostAdapter();
    const executor = new ToolExecutor(host);
    const badExpand = await executor.execute({
      name: "range.read",
      arguments: { sheetName: "Sheet1", range: "A1", expand: "usedRange" },
    });
    expect(badExpand.ok).toBe(false);
    if (!badExpand.ok) expect(badExpand.error).toContain("expand");

    const badOp = await executor.execute({
      name: "sheet.operation",
      arguments: { operation: "merge", sheetName: "Sheet1" },
    });
    expect(badOp.ok).toBe(false);
    if (!badOp.ok) expect(badOp.error).toContain("operation");

    const badPos = await executor.execute({
      name: "sheet.operation",
      arguments: { operation: "move", sheetName: "Sheet1", position: 0 },
    });
    expect(badPos.ok).toBe(false);
    if (!badPos.ok) expect(badPos.error).toContain("1-based");
  });
});

