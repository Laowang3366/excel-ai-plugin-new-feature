import { beforeEach, describe, expect, test, vi } from "vitest";
import { executePowerShell } from "../../../automation/powershell";
import { executeSmart } from "../../../automation/scriptEngine";
import {
  clearRangeOperation,
  getSelectionOperation,
  readRangeOperation,
  writeRangeOperation,
  type RangeOperationDeps,
} from "./rangeOperations";

vi.mock("../../../automation/powershell", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../../automation/powershell")>();
  return {
    ...actual,
    executePowerShell: vi.fn(),
  };
});

vi.mock("../../../automation/scriptEngine", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../../automation/scriptEngine")>();
  return {
    ...actual,
    executeSmart: vi.fn(),
  };
});

const executePowerShellMock = vi.mocked(executePowerShell);
const executeSmartMock = vi.mocked(executeSmart);

function connectedDeps(): RangeOperationDeps {
  return {
    ensureConnected: vi.fn(async () => "excel" as const),
    getProgId: () => "Excel.Application",
  };
}

describe("rangeOperations read performance path", () => {
  beforeEach(() => {
    executePowerShellMock.mockReset();
    executeSmartMock.mockReset();
    executeSmartMock.mockResolvedValue({ result: "OK", engine: "python" });
  });

  test("range.read uses cached connection and bulk Value2 reads", async () => {
    executePowerShellMock.mockResolvedValueOnce(JSON.stringify({
      values: [["A", "B"], [1, 2]],
      address: "A1:B2",
      expanded: false,
      expandMode: "none",
    }));
    const deps = connectedDeps();

    const values = await readRangeOperation(deps, "Sheet1", "A1:B2");

    expect(values).toEqual({
      values: [["A", "B"], [1, 2]],
      address: "A1:B2",
      expanded: false,
      expandMode: "none",
    });
    expect(deps.ensureConnected).toHaveBeenCalledWith(0);
    const script = executePowerShellMock.mock.calls[0][0];
    expect(script).toContain("$readRange.Value2");
    expect(script).toContain("[System.Collections.Generic.List[object]]::new()");
    expect(script).not.toContain("$rows +=");
    expect(script).not.toContain("$row +=");
  });

  test("range.read can expand dynamic array spill ranges for validation", async () => {
    executePowerShellMock.mockResolvedValueOnce(JSON.stringify({
      values: [["A"], ["B"], ["C"]],
      address: "H2:H4",
      expanded: true,
      expandMode: "spill",
    }));
    const deps = connectedDeps();

    const result = await readRangeOperation(deps, "Sheet1", "H2", "spill");

    expect(result).toEqual({
      values: [["A"], ["B"], ["C"]],
      address: "H2:H4",
      expanded: true,
      expandMode: "spill",
    });
    const script = executePowerShellMock.mock.calls[0][0];
    expect(script).toContain("SpillingToRange");
    expect(script).toContain("CurrentArray");
    expect(script).toContain("$readRange.Value2");
  });

  test("selection.get reads address and values in one PowerShell call", async () => {
    executePowerShellMock.mockResolvedValueOnce(JSON.stringify({
      address: "A1:B2",
      sheetName: "Sheet1",
      values: [["A", "B"], [1, 2]],
    }));
    const deps = connectedDeps();
    const readRange = vi.fn();

    const selection = await getSelectionOperation(deps, readRange);

    expect(selection).toEqual({
      address: "A1:B2",
      sheetName: "Sheet1",
      values: [["A", "B"], [1, 2]],
    });
    expect(deps.ensureConnected).toHaveBeenCalledWith(0);
    expect(readRange).not.toHaveBeenCalled();
    const script = executePowerShellMock.mock.calls[0][0];
    expect(script).toContain("$sel.Value2");
    expect(script).toContain("[System.Collections.Generic.List[object]]::new()");
    expect(script).not.toContain("$sel.Cells.Item");
    expect(script).not.toContain("$rows +=");
    expect(script).not.toContain("$row +=");
  });

  test("range write and clear keep executeSmart on the legacy non-Python-first path", async () => {
    const deps = connectedDeps();

    await writeRangeOperation(deps, "Sheet1", "A1", [["A"]]);
    await clearRangeOperation(deps, "Sheet1", "A1");

    expect(executeSmartMock).toHaveBeenCalledTimes(2);
    expect(executeSmartMock.mock.calls[0][4]).toEqual({ preferPython: false });
    expect(executeSmartMock.mock.calls[1][4]).toEqual({ preferPython: false });
    const [pythonScript, jscriptScript, powershellScript] = executeSmartMock.mock.calls[0];
    expect(pythonScript).toContain("rng.cells.item(r, c).value = val");
    expect(jscriptScript).toContain("rng.Cells.Item(1, 1) =");
    expect(powershellScript).toContain("$startRange.Cells.Item(1, 1) =");
    expect(jscriptScript).not.toContain(".Value2 =");
    expect(powershellScript).not.toContain(".Value2 =");
  });

  test("range write does not special-case formulas or dynamic arrays", async () => {
    const deps = connectedDeps();

    await writeRangeOperation(deps, "Sheet1", "A1", [
      ["=FILTER(A:A,A:A>0)"],
      [""],
      ["plain value"],
    ]);

    expect(executeSmartMock).toHaveBeenCalledTimes(1);
    const [pythonScript, jscriptScript, powershellScript] = executeSmartMock.mock.calls[0];
    const combinedScript = `${pythonScript}\n${jscriptScript}\n${powershellScript}`;
    expect(executeSmartMock.mock.calls[0][4]).toEqual({ preferPython: false });
    expect(pythonScript).toContain("rng.cells.item(r, c).value = val");
    expect(jscriptScript).toContain("rng.Cells.Item(1, 1) =");
    expect(powershellScript).toContain("$startRange.Cells.Item(1, 1) =");
    expect(combinedScript).not.toContain("Formula2");
    expect(combinedScript).not.toContain(".Formula");
    expect(combinedScript).not.toContain(".Value2 =");
    expect(combinedScript).not.toContain("DYNAMIC_ARRAY_FORMULA");
  });
});
