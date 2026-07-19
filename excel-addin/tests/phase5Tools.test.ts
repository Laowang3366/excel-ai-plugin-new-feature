import { describe, expect, it } from "vitest";
import { ToolExecutor, TOOL_DEFINITIONS } from "../shared/tools";
import { MockHostAdapter } from "./mockHost";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";

describe("phase5 tools", () => {
  it("registers conditionalFormat and dataValidation tools", () => {
    const names = TOOL_DEFINITIONS.map((t) => t.name);
    expect(names).toContain("conditionalFormat.add");
    expect(names).toContain("dataValidation.write");
  });

  it("adds cellValue CF and list DV via mock host", async () => {
    const host = new MockHostAdapter();
    const executor = new ToolExecutor(host);
    const added = await executor.execute({
      name: "conditionalFormat.add",
      arguments: {
        sheetName: "Sheet1",
        range: "A1:A10",
        rule: {
          kind: "cellValue",
          operator: "greaterThan",
          formula1: "10",
          fillColor: "#FF0000",
        },
      },
    });
    expect(added.ok).toBe(true);
    if (added.ok) {
      const id = (added.data as { id: string }).id;
      const listed = await executor.execute({
        name: "conditionalFormat.list",
        arguments: { sheetName: "Sheet1", range: "A1:A10" },
      });
      expect(listed.ok).toBe(true);
      await executor.execute({
        name: "conditionalFormat.delete",
        arguments: { sheetName: "Sheet1", range: "A1:A10", id },
      });
    }

    const dv = await executor.execute({
      name: "dataValidation.write",
      arguments: {
        sheetName: "Sheet1",
        range: "B1:B5",
        rule: { type: "list", listValues: ["A", "B", "C"] },
      },
    });
    expect(dv.ok).toBe(true);
    const read = await executor.execute({
      name: "dataValidation.read",
      arguments: { sheetName: "Sheet1", range: "B1:B5" },
    });
    expect(read.ok).toBe(true);
    if (read.ok) {
      expect((read.data as { rule: { type: string } }).rule?.type).toBe("list");
    }
    expect(
      (
        await executor.execute({
          name: "dataValidation.clear",
          arguments: { sheetName: "Sheet1", range: "B1:B5" },
        })
      ).ok,
    ).toBe(true);
  });

  it("rejects unknown fields, missing operator, bad listValues, unimplemented showError", async () => {
    const host = new MockHostAdapter();
    const executor = new ToolExecutor(host);
    const bad = await executor.execute({
      name: "conditionalFormat.add",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: { kind: "cellValue", formula1: "1", operator: "startsWith", extra: true },
      },
    });
    expect(bad.ok).toBe(false);

    const missingOp = await executor.execute({
      name: "conditionalFormat.add",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: { kind: "cellValue", formula1: "1" },
      },
    });
    expect(missingOp.ok).toBe(false);

    const badDv = await executor.execute({
      name: "dataValidation.write",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: { type: "date", formula1: "1" },
      },
    });
    expect(badDv.ok).toBe(false);

    const badList = await executor.execute({
      name: "dataValidation.write",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: { type: "list", listValues: [1, "x"] },
      },
    });
    expect(badList.ok).toBe(false);

    const showError = await executor.execute({
      name: "dataValidation.write",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: { type: "list", listValues: ["A"], showError: true },
      },
    });
    expect(showError.ok).toBe(false);

    const wholeMissingOp = await executor.execute({
      name: "dataValidation.write",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: { type: "wholeNumber", formula1: "1" },
      },
    });
    expect(wholeMissingOp.ok).toBe(false);
  });

  it("WPS returns unsupported for CF and DV", async () => {
    const host = new WpsJsaAdapter();
    const executor = new ToolExecutor(host);
    for (const call of [
      {
        name: "conditionalFormat.list" as const,
        arguments: { sheetName: "Sheet1", range: "A1" },
      },
      {
        name: "dataValidation.read" as const,
        arguments: { sheetName: "Sheet1", range: "A1" },
      },
    ]) {
      const result = await executor.execute(call);
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.unsupported).toBe(true);
    }
  });
});
