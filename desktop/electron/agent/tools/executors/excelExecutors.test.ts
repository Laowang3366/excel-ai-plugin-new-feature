import { describe, expect, it, vi } from "vitest";

import type { ToolExecutor } from "../../shared/types";
import type {
  ExcelUiBridge,
  ExcelVbaBridge,
  ExcelWorkbookBridge,
  WpsJsaBridge,
} from "../contracts/excel";
import { addExcelExecutors } from "./excelExecutors";

function createExcelExecutors(
  overrides: Partial<{
    workbookBridge: ExcelWorkbookBridge;
    vbaBridge: ExcelVbaBridge;
    jsaBridge: WpsJsaBridge;
    uiBridge: ExcelUiBridge;
  }> = {},
): Map<string, ToolExecutor> {
  const target = new Map<string, ToolExecutor>();
  addExcelExecutors(target, {
    workbookBridge: {
      readRange: vi.fn(),
      writeRange: vi.fn(),
      openWorkbook: vi.fn(),
    } as unknown as ExcelWorkbookBridge,
    vbaBridge: {} as ExcelVbaBridge,
    jsaBridge: {} as WpsJsaBridge,
    uiBridge: {} as ExcelUiBridge,
    ...overrides,
  });
  return target;
}

describe("addExcelExecutors", () => {
  it("keeps the original registration order after macro and UI extraction", () => {
    const target = createExcelExecutors();

    expect([...target.keys()]).toEqual([
      "workbook.inspect",
      "range.read",
      "range.write",
      "range.clear",
      "selection.get",
      "formula.context",
      "macro.detect",
      "macro.run",
      "macro.write",
      "sheet.operation",
      "ui.addControl",
      "ui.removeControl",
      "ui.listControls",
      "ui.createForm",
      "ui.addMenu",
      "workbook.open",
      "workbook.create",
      "workbook.save",
      "workbook.switch",
    ]);
  });

  it("does not expose the removed external script or language-specific VBA tools", () => {
    const target = createExcelExecutors();

    expect(target.has("script.execute")).toBe(false);
    expect(target.has("script.detect")).toBe(false);
    expect(target.has("vba.writeModule")).toBe(false);
    expect(target.has("vba.runMacro")).toBe(false);
    expect(target.has("macro.write")).toBe(true);
  });

  it("installs and saves a VBA module through the unified macro tool", async () => {
    const writeResult = {
      moduleName: "UnitButtons",
      created: true,
      lineCount: 3,
      sourceVerified: true as const,
      compileVerified: true as const,
      entryPoint: "UnitButtonClick",
      entryPointVerified: true,
      saved: true,
      workbookName: "Book-macro.xlsm",
      workbookPath: "D:\\docs\\Book-macro.xlsm",
      host: "wps" as const,
    };
    const vbaBridge = {
      writeModule: vi.fn(async () => writeResult),
    } as unknown as ExcelVbaBridge;
    const target = createExcelExecutors({ vbaBridge });

    const result = await target.get("macro.write")!.execute({
      language: "vba",
      moduleName: "UnitButtons",
      code: "Public Sub UnitButtonClick()\nEnd Sub",
      entryPoint: "UnitButtonClick",
    });

    expect(result).toEqual({
      success: true,
      data: { language: "vba", ...writeResult },
    });
    expect(vbaBridge.writeModule).toHaveBeenCalledWith(
      "UnitButtons",
      "Public Sub UnitButtonClick()\nEnd Sub",
      { entryPoint: "UnitButtonClick", save: true, saveAsPath: undefined },
    );
  });

  it("requires an entry point before installing a VBA module", async () => {
    const vbaBridge = { writeModule: vi.fn() } as unknown as ExcelVbaBridge;
    const target = createExcelExecutors({ vbaBridge });

    const result = await target.get("macro.write")!.execute({
      language: "vba",
      moduleName: "UnitButtons",
      code: "Public Sub UnitButtonClick()\nEnd Sub",
    });

    expect(result).toEqual({
      success: false,
      error: "缺少必填参数: entryPoint",
    });
    expect(vbaBridge.writeModule).not.toHaveBeenCalled();
  });

  it("detects only macro languages that live inside the workbook host", async () => {
    const workbookBridge = {
      getHostInfo: vi.fn(async () => ({ host: "wps" as const, version: "12" })),
    } as unknown as ExcelWorkbookBridge;
    const vbaBridge = {
      detectCapabilities: vi.fn(async () => ({
        supported: true,
        host: "wps" as const,
      })),
    } as unknown as ExcelVbaBridge;
    const jsaBridge = {
      detectCapabilities: vi.fn(async () => ({
        language: "javascript" as const,
        supported: true,
        ready: true,
        internal: true as const,
        engine: "WPS JSA" as const,
      })),
    } as unknown as WpsJsaBridge;
    const target = createExcelExecutors({
      workbookBridge,
      vbaBridge,
      jsaBridge,
    });

    const result = await target.get("macro.detect")!.execute({});

    expect(result).toMatchObject({
      success: true,
      data: {
        host: "wps",
        recommended: "vba",
        available: [
          { language: "vba", ready: true, internal: true, engine: "VBA" },
          {
            language: "javascript",
            ready: true,
            internal: true,
            engine: "WPS JSA",
          },
        ],
      },
    });
  });

  it("writes JavaScript only through the WPS internal JSA bridge", async () => {
    const writeResult = {
      language: "javascript" as const,
      componentName: "Module1",
      lineCount: 2,
      sourceVerified: true as const,
      entryPoint: "main",
      entryPointVerified: true,
      saved: true,
      host: "wps" as const,
    };
    const jsaBridge = {
      writeCode: vi.fn(async () => writeResult),
    } as unknown as WpsJsaBridge;
    const target = createExcelExecutors({ jsaBridge });

    const result = await target.get("macro.write")!.execute({
      language: "javascript",
      code: "function main() {}",
      entryPoint: "main",
    });

    expect(result).toEqual({ success: true, data: writeResult });
    expect(jsaBridge.writeCode).toHaveBeenCalledWith("function main() {}", {
      entryPoint: "main",
      save: true,
    });
  });

  it("does not claim an unverified WPS JSA remote-run capability", async () => {
    const target = createExcelExecutors();

    const result = await target.get("macro.run")!.execute({
      language: "javascript",
      macroName: "main",
    });

    expect(result).toEqual({
      success: false,
      error: "macro.run 当前仅支持 vba；WPS JSA 只提供写入和回读校验",
    });
  });

  it("omits host version from model-facing workbook inspection metadata", async () => {
    const workbookBridge = {
      inspectWorkbook: vi.fn(async () => ({
        name: "WPS 表格",
        host: "wps",
        version: "12.0",
        workbooks: [{ name: "demo.xlsx", sheets: [] }],
      })),
    } as unknown as ExcelWorkbookBridge;
    const target = createExcelExecutors({ workbookBridge });

    const result = await target.get("workbook.inspect")!.execute({});

    expect(result).toEqual({
      success: true,
      data: {
        name: "WPS 表格",
        host: "wps",
        workbooks: [{ name: "demo.xlsx", sheets: [] }],
        formulaDialect: {
          regexFunction: "REGEXP",
          guidance:
            "WPS 正则提取使用 REGEXP；不要使用 Excel 方言的 REGEXEXTRACT/REGEXREPLACE/REGEXTEST 函数名",
        },
      },
    });
  });

  it("keeps range.read backward compatible by returning values by default", async () => {
    const workbookBridge = {
      readRange: vi.fn(async () => ({
        values: [["A"], ["B"]],
        address: "A1:A2",
        expanded: false,
        expandMode: "none",
      })),
    } as unknown as ExcelWorkbookBridge;
    const target = createExcelExecutors({ workbookBridge });

    const result = await target.get("range.read")!.execute({
      sheetName: "Sheet1",
      range: "A1:A2",
    });

    expect(result).toEqual({ success: true, data: [["A"], ["B"]] });
    expect(workbookBridge.readRange).toHaveBeenCalledWith("Sheet1", "A1:A2", "none");
  });

  it("automatically reads the full spill when a single-cell read omits expand", async () => {
    const workbookBridge = {
      readRange: vi.fn(async () => ({
        values: [[6], [7], [8]],
        address: "J20:J22",
        expanded: true,
        expandMode: "spill" as const,
      })),
    } as unknown as ExcelWorkbookBridge;
    const target = createExcelExecutors({ workbookBridge });

    const result = await target.get("range.read")!.execute({
      sheetName: "Sheet1",
      range: "J20",
    });

    expect(result).toEqual({ success: true, data: [[6], [7], [8]] });
    expect(workbookBridge.readRange).toHaveBeenCalledWith("Sheet1", "J20", "spill");
  });

  it("respects an explicit none mode for single-cell reads", async () => {
    const workbookBridge = {
      readRange: vi.fn(async () => ({
        values: [[6]],
        address: "J20",
        expanded: false,
        expandMode: "none" as const,
      })),
    } as unknown as ExcelWorkbookBridge;
    const target = createExcelExecutors({ workbookBridge });

    const result = await target.get("range.read")!.execute({
      sheetName: "Sheet1",
      range: "$J$20",
      expand: "none",
    });

    expect(result).toEqual({ success: true, data: [[6]] });
    expect(workbookBridge.readRange).toHaveBeenCalledWith("Sheet1", "$J$20", "none");
  });

  it("returns expanded range metadata when range.read uses spill mode", async () => {
    const expanded = {
      values: [["A"], ["B"]],
      address: "H2:H3",
      expanded: true,
      expandMode: "spill" as const,
    };
    const workbookBridge = {
      readRange: vi.fn(async () => expanded),
    } as unknown as ExcelWorkbookBridge;
    const target = createExcelExecutors({ workbookBridge });

    const result = await target.get("range.read")!.execute({
      sheetName: "Sheet1",
      range: "H2",
      expand: "spill",
    });

    expect(result).toEqual({ success: true, data: expanded });
    expect(workbookBridge.readRange).toHaveBeenCalledWith("Sheet1", "H2", "spill");
  });

  it("rejects invalid range.read expand modes", async () => {
    const target = createExcelExecutors();

    const result = await target.get("range.read")!.execute({
      sheetName: "Sheet1",
      range: "A1",
      expand: "usedRange",
    });

    expect(result).toEqual({
      success: false,
      error: "参数 expand 必须是 none、spill、currentArray 或 currentRegion",
    });
  });

  it("validates range.write values before writing", async () => {
    const workbookBridge = {
      writeRange: vi.fn(),
    } as unknown as ExcelWorkbookBridge;
    const target = createExcelExecutors({ workbookBridge });

    const result = await target.get("range.write")!.execute({
      sheetName: "Sheet1",
      range: "A1",
    });

    expect(result).toEqual({
      success: false,
      error: "缺少必填参数: values",
    });
    expect(workbookBridge.writeRange).not.toHaveBeenCalled();
  });

  it("accepts range.write values supplied as a JSON string", async () => {
    const writeResult = {
      written: 2,
      dynamicCells: 0,
      arrayCells: 0,
      plainCells: 0,
    };
    const workbookBridge = {
      writeRange: vi.fn(async () => writeResult),
    } as unknown as ExcelWorkbookBridge;
    const target = createExcelExecutors({ workbookBridge });

    const result = await target.get("range.write")!.execute({
      sheetName: "Sheet1",
      range: "A1:B1",
      values: "[[1,2]]",
    });

    expect(result).toEqual({ success: true, data: writeResult });
    expect(workbookBridge.writeRange).toHaveBeenCalledWith("Sheet1", "A1:B1", [[1, 2]], {
      legacyCse: false,
    });
  });

  it("passes legacy CSE intent only when explicitly requested", async () => {
    const writeResult = {
      written: 1,
      dynamicCells: 0,
      arrayCells: 1,
      plainCells: 0,
    };
    const workbookBridge = {
      writeRange: vi.fn(async () => writeResult),
    } as unknown as ExcelWorkbookBridge;
    const target = createExcelExecutors({ workbookBridge });

    const result = await target.get("range.write")!.execute({
      sheetName: "Sheet1",
      range: "A1",
      values: [["=SUM(A1:A10)"]],
      legacyCse: true,
    });

    expect(result).toEqual({ success: true, data: writeResult });
    expect(workbookBridge.writeRange).toHaveBeenCalledWith("Sheet1", "A1", [["=SUM(A1:A10)"]], {
      legacyCse: true,
    });
  });

  it("rejects invalid range.write JSON string values", async () => {
    const workbookBridge = {
      writeRange: vi.fn(),
    } as unknown as ExcelWorkbookBridge;
    const target = createExcelExecutors({ workbookBridge });

    const result = await target.get("range.write")!.execute({
      sheetName: "Sheet1",
      range: "A1",
      values: "{bad-json",
    });

    expect(result).toEqual({
      success: false,
      error: "参数 values 应为数组，但收到了字符串且无法解析为 JSON",
    });
    expect(workbookBridge.writeRange).not.toHaveBeenCalled();
  });

  it("rejects WPS-hostile formula typography before writing", async () => {
    const workbookBridge = {
      writeRange: vi.fn(),
    } as unknown as ExcelWorkbookBridge;
    const target = createExcelExecutors({ workbookBridge });

    const result = await target.get("range.write")!.execute({
      sheetName: "Sheet1",
      range: "A1",
      values: [["=LET(t,A1,—LEFT(t,2))"]],
    });

    expect(result.success).toBe(false);
    expect(result.error).toContain("WPS 公式解析风险");
    expect(result.error).toContain("—");
    expect(result.error).toContain("ASCII 减号");
    expect(workbookBridge.writeRange).not.toHaveBeenCalled();
  });

  it("validates workbook.open filePath before opening", async () => {
    const workbookBridge = {
      openWorkbook: vi.fn(),
    } as unknown as ExcelWorkbookBridge;
    const target = createExcelExecutors({ workbookBridge });

    const result = await target.get("workbook.open")!.execute({});

    expect(result).toEqual({
      success: false,
      error: "缺少必填参数: filePath",
    });
    expect(workbookBridge.openWorkbook).not.toHaveBeenCalled();
  });
});
