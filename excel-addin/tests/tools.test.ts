import { describe, expect, it } from "vitest";
import { ToolExecutor, TOOL_DEFINITIONS } from "../shared/tools";
import { MockHostAdapter } from "./mockHost";

describe("ToolExecutor", () => {
  it("exposes first-batch tool definitions", () => {
    const names = TOOL_DEFINITIONS.map((tool) => tool.name);
    expect(names).toContain("selection.get");
    expect(names).toContain("range.write");
    expect(names).toContain("formula.write");
    expect(names).toContain("sheet.list");
  });

  it("reads selection from host", async () => {
    const host = new MockHostAdapter();
    const executor = new ToolExecutor(host);
    const result = await executor.execute({ name: "selection.get", arguments: {} });
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.data).toMatchObject({ sheetName: "Sheet1" });
    }
  });

  it("writes range and verifies with read-after-write", async () => {
    const host = new MockHostAdapter();
    const executor = new ToolExecutor(host);
    const result = await executor.execute({
      name: "range.write",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        values: [["ok"]],
        verify: true,
      },
    });
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.verification).toMatchObject({ ok: true });
    }
  });

  it("writes formula via writeFormulas primitive with verify", async () => {
    const host = new MockHostAdapter();
    const executor = new ToolExecutor(host);
    const result = await executor.execute({
      name: "formula.write",
      arguments: {
        sheetName: "Sheet1",
        range: "B1",
        formula: "SUM(1,2)",
        verify: true,
      },
    });
    expect(result.ok).toBe(true);
    if (result.ok) {
      const data = result.data as { formulas: string[][] };
      expect(data.formulas[0][0]).toBe("=SUM(1,2)");
    }
    // Values-only write must not populate formulas (formula path is separate).
    await executor.execute({
      name: "range.write",
      arguments: {
        sheetName: "Sheet1",
        range: "C1",
        values: [["=not-a-formula-path"]],
        verify: false,
      },
    });
    const read = await host.readRange("Sheet1", "C1");
    expect(read.ok && read.data.formulas[0][0]).toBe("");
  });

  it("lists and mutates sheets", async () => {
    const host = new MockHostAdapter();
    const executor = new ToolExecutor(host);
    await executor.execute({ name: "sheet.add", arguments: { sheetName: "Data" } });
    await executor.execute({
      name: "sheet.rename",
      arguments: { sheetName: "Data", newName: "Metrics" },
    });
    const listed = await executor.execute({ name: "sheet.list", arguments: {} });
    expect(listed.ok).toBe(true);
    if (listed.ok) {
      const names = (listed.data as { name: string }[]).map((sheet) => sheet.name);
      expect(names).toContain("Metrics");
    }
  });
});
