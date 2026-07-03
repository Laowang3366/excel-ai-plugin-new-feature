import { beforeEach, describe, expect, test, vi } from "vitest";
import { executePowerShell } from "../../../automation/powershell";
import { getFormulaContextOperation, type FormulaOperationDeps } from "./formulaOperations";

vi.mock("../../../automation/powershell", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../../automation/powershell")>();
  return {
    ...actual,
    executePowerShell: vi.fn(),
  };
});

const executePowerShellMock = vi.mocked(executePowerShell);

function connectedDeps(): FormulaOperationDeps {
  return {
    ensureConnected: vi.fn(async () => "excel" as const),
    getProgId: () => "Excel.Application",
  };
}

describe("formulaOperations", () => {
  beforeEach(() => {
    executePowerShellMock.mockReset();
  });

  test("formula.context uses Formula2 with Formula fallback and Value2 reads", async () => {
    executePowerShellMock.mockResolvedValueOnce(JSON.stringify([
      { cell: "B2", formula: "=SUM(A1:A2)", value: "3" },
    ]));
    const deps = connectedDeps();

    const result = await getFormulaContextOperation(deps, "Sheet1", "A1:C10");

    expect(result).toEqual([
      { cell: "B2", formula: "=SUM(A1:A2)", value: "3" },
    ]);
    expect(deps.ensureConnected).toHaveBeenCalledWith(0);
    const script = executePowerShellMock.mock.calls[0][0];
    expect(script).toContain("$formulaValues = $range.Formula2");
    expect(script).toContain("$formulaValues = $range.Formula");
    expect(script).toContain("$cellValues = $range.Value2");
    expect(script).toContain("[System.Collections.Generic.List[object]]::new()");
    expect(script).not.toContain("$range.Cells.Item");
  });
});
