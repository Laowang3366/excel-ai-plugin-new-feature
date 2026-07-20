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
          operator: "greaterThanOrEqualTo",
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
        rule: { type: "decimal", operator: "between", formula1: "0", formula2: "1" },
      },
    });
    expect(dv.ok).toBe(true);
    const read = await executor.execute({
      name: "dataValidation.read",
      arguments: { sheetName: "Sheet1", range: "B1:B5" },
    });
    expect(read.ok).toBe(true);
    if (read.ok) {
      expect((read.data as { rule: { type: string } }).rule?.type).toBe("decimal");
    }
  });

  it("strictly rejects bad CF/DV rule shapes", async () => {
    const host = new MockHostAdapter();
    const executor = new ToolExecutor(host);

    const badColor = await executor.execute({
      name: "conditionalFormat.add",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: {
          kind: "cellValue",
          operator: "greaterThan",
          formula1: "1",
          fillColor: 123 as unknown as string,
        },
      },
    });
    expect(badColor.ok).toBe(false);

    const formula2OnGt = await executor.execute({
      name: "conditionalFormat.add",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: {
          kind: "cellValue",
          operator: "greaterThan",
          formula1: "1",
          formula2: "2",
        },
      },
    });
    expect(formula2OnGt.ok).toBe(false);

    const bothList = await executor.execute({
      name: "dataValidation.write",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: { type: "list", listValues: ["A"], formula1: "Sheet1!A1:A2" },
      },
    });
    expect(bothList.ok).toBe(false);

    const showError = await executor.execute({
      name: "dataValidation.write",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: { type: "list", listValues: ["A"], showError: true },
      },
    });
    expect(showError.ok).toBe(false);

    const emptyFormula = await executor.execute({
      name: "conditionalFormat.add",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: { kind: "custom", formula: "   " },
      },
    });
    expect(emptyFormula.ok).toBe(false);

    const emptyColor = await executor.execute({
      name: "conditionalFormat.add",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: {
          kind: "cellValue",
          operator: "greaterThan",
          formula1: "1",
          fillColor: "",
        },
      },
    });
    expect(emptyColor.ok).toBe(false);

    const emptyFormula2 = await executor.execute({
      name: "conditionalFormat.add",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: {
          kind: "cellValue",
          operator: "greaterThan",
          formula1: "1",
          formula2: "",
        },
      },
    });
    expect(emptyFormula2.ok).toBe(false);

    const badAllowBlank = await executor.execute({
      name: "dataValidation.write",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: { type: "list", listValues: ["A"], allowBlank: "yes" },
      },
    });
    expect(badAllowBlank.ok).toBe(false);

    const tooMany = await executor.execute({
      name: "dataValidation.write",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: {
          type: "list",
          listValues: Array.from({ length: 1001 }, (_, i) => `v${i}`),
        },
      },
    });
    expect(tooMany.ok).toBe(false);

    const tooLongInline = await executor.execute({
      name: "dataValidation.write",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: {
          type: "list",
          listValues: [ "x".repeat(200), "y".repeat(200) ],
        },
      },
    });
    expect(tooLongInline.ok).toBe(false);
  });

  it("WPS returns unsupported for all six CF/DV tools", async () => {
    const host = new WpsJsaAdapter();
    const executor = new ToolExecutor(host);
    for (const call of [
      { name: "conditionalFormat.list" as const, arguments: { sheetName: "Sheet1", range: "A1" } },
      {
        name: "conditionalFormat.add" as const,
        arguments: {
          sheetName: "Sheet1",
          range: "A1",
          rule: { kind: "custom", formula: "=TRUE" },
        },
      },
      {
        name: "conditionalFormat.delete" as const,
        arguments: { sheetName: "Sheet1", range: "A1", id: "x" },
      },
      { name: "dataValidation.read" as const, arguments: { sheetName: "Sheet1", range: "A1" } },
      {
        name: "dataValidation.write" as const,
        arguments: {
          sheetName: "Sheet1",
          range: "A1",
          rule: { type: "list", listValues: ["A"] },
        },
      },
      {
        name: "dataValidation.clear" as const,
        arguments: { sheetName: "Sheet1", range: "A1" },
      },
    ]) {
      const result = await executor.execute(call);
      expect(result.ok, call.name).toBe(false);
      if (!result.ok) expect(result.unsupported, call.name).toBe(true);
    }
  });
});
