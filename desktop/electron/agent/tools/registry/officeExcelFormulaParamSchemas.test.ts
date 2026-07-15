import { describe, expect, it } from "vitest";

import { ALL_TOOL_DEFINITIONS } from "./toolDefinitions";
import { parseAndValidateToolArguments } from "./toolSchema";

function parameters(name: string) {
  const definition = ALL_TOOL_DEFINITIONS.find((tool) => tool.name === name);
  if (!definition) throw new Error(`missing tool definition: ${name}`);
  return definition.parameters;
}

describe("Excel formula governance parameter schemas", () => {
  it("accepts repair parameters in apply and workflow", () => {
    const step = {
      app: "excel",
      action: "edit",
      operation: "repairFormulaReferences",
      filePath: "C:/book.xlsx",
      target: "range:Summary!D2:D100",
      params: {
        host: "excel",
        scope: "target",
        replacements: [
          { find: "#REF!$A$2", replace: "Source!$A$2" },
          { find: "#REF!", replace: "Archive" },
        ],
      },
    };

    expect(
      parseAndValidateToolArguments(JSON.stringify(step), parameters("office.action.apply")).error,
    ).toBeUndefined();
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({ steps: [step] }),
        parameters("office.workflow.run"),
      ).error,
    ).toBeUndefined();
  });

  it("rejects missing, unknown, and malformed repair parameters", () => {
    const base = {
      app: "excel",
      action: "edit",
      operation: "repairFormulaReferences",
      filePath: "C:/book.xlsx",
    };
    const schema = parameters("office.action.apply");

    expect(parseAndValidateToolArguments(JSON.stringify(base), schema).error).toContain("params");
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({ ...base, params: { replacements: [] } }),
        schema,
      ).error,
    ).toContain("replacements");
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({
          ...base,
          params: { replacements: [{ find: "#REF!", regex: true }] },
        }),
        schema,
      ).error,
    ).toContain("regex");
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({
          ...base,
          params: { replacements: [{ find: "#REF!" }], copyFromNeighbors: true },
        }),
        schema,
      ).error,
    ).toContain("copyFromNeighbors");
  });

  it("constrains backup, restore, and protection operations", () => {
    const apply = parameters("office.action.apply");
    const validOperations = [
      {
        operation: "convertFormulasToValues",
        params: { scope: "sheet", createBackup: true, backupId: "batch-2026-07" },
      },
      {
        operation: "restoreFormulas",
        params: { backupId: "batch-2026-07", removeAfterRestore: true },
      },
      {
        operation: "manageFormulaProtection",
        params: { scope: "workbook", command: "lock", protectSheet: true, password: "secret" },
      },
    ];

    for (const item of validOperations) {
      const args = {
        app: "excel",
        action: "edit",
        operation: item.operation,
        filePath: "C:/book.xlsx",
        params: item.params,
      };
      expect(parseAndValidateToolArguments(JSON.stringify(args), apply).error).toBeUndefined();
    }

    const invalid = {
      app: "excel",
      action: "edit",
      operation: "manageFormulaProtection",
      filePath: "C:/book.xlsx",
      params: { command: "encrypt" },
    };
    expect(parseAndValidateToolArguments(JSON.stringify(invalid), apply).error).toContain(
      "command",
    );
  });

  it("uses strict read-only schemas for inspect and validate", () => {
    const valid = {
      app: "excel",
      operation: "inspectFormulaDependencies",
      filePath: "C:/book.xlsx",
      params: { host: "wps", scope: "workbook" },
    };

    for (const name of ["office.action.inspect", "office.action.validate"]) {
      const schema = parameters(name);
      expect(parseAndValidateToolArguments(JSON.stringify(valid), schema).error).toBeUndefined();
      expect(
        parseAndValidateToolArguments(
          JSON.stringify({ ...valid, params: { ...valid.params, maxExpandedRangeCells: 500 } }),
          schema,
        ).error,
      ).toContain("maxExpandedRangeCells");
    }
  });
});
