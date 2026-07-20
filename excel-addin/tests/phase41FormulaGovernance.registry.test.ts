import { describe, expect, it } from "vitest";
import { ToolExecutor } from "../shared/tools/executor";
import { TOOL_DEFINITIONS } from "../shared/tools/definitions";
import { dispositionForRisk } from "../shared/agentChat/approvalPolicy";
import { CHAT_READONLY_TOOL_ALLOWLIST } from "../shared/agentChat/chatReadOnlyTools";
import { MockHostAdapter } from "./mockHost";

describe("phase41 formula governance registry", () => {
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
    const restore = TOOL_DEFINITIONS.find((t) => t.name === "formula.backups.restore");
    expect(restore?.parameters.properties).toMatchObject({
      backupId: { type: "string" },
      removeAfterRestore: { type: "boolean" },
    });
  });

  it("executor rejects unknown fields and incomplete repair without write", async () => {
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

  it("mock convert backs up formula text then restore / removeAfterRestore", async () => {
    const host = new MockHostAdapter();
    host.cells.set("Sheet1!A1", { values: [[42]], formulas: [["=6*7"]] });
    host.cells.set("Sheet1!B1", { values: [[1]], formulas: [["=A1"]] });
    const ex = new ToolExecutor(host);

    const converted = await ex.execute({
      name: "formula.convertToValues",
      arguments: { scope: "target", sheetName: "Sheet1", range: "A1", backupId: "b1" },
    });
    expect(converted.ok).toBe(true);
    expect(host.formulaBackupRows.some((r) => r.backupId === "b1" && r.formula === "=6*7")).toBe(
      true,
    );

    // second backup id
    await ex.execute({
      name: "formula.convertToValues",
      arguments: { scope: "target", sheetName: "Sheet1", range: "B1", backupId: "b2" },
    });

    const restored = await ex.execute({
      name: "formula.backups.restore",
      arguments: { backupId: "b1", removeAfterRestore: true },
    });
    expect(restored.ok).toBe(true);
    expect(host.cells.get("Sheet1!A1")?.formulas[0]?.[0]).toBe("=6*7");
    expect(host.formulaBackupRows.some((r) => r.backupId === "b1")).toBe(false);
    expect(host.formulaBackupRows.some((r) => r.backupId === "b2")).toBe(true);
  });
});
