import { describe, expect, it, vi } from "vitest";

import type { ToolExecutor } from "../../shared/types";
import type {
  ExcelScriptBridge,
  ExcelUiBridge,
  ExcelVbaBridge,
  ExcelWorkbookBridge,
} from "../contracts/excel";
import { addExcelExecutors } from "./excelExecutors";

function createExcelExecutors(overrides: Partial<{
  workbookBridge: ExcelWorkbookBridge;
  vbaBridge: ExcelVbaBridge;
  scriptBridge: ExcelScriptBridge;
  uiBridge: ExcelUiBridge;
}> = {}): Map<string, ToolExecutor> {
  const target = new Map<string, ToolExecutor>();
  addExcelExecutors(target, {
    workbookBridge: {
      readRange: vi.fn(),
      writeRange: vi.fn(),
      openWorkbook: vi.fn(),
    } as unknown as ExcelWorkbookBridge,
    vbaBridge: {} as ExcelVbaBridge,
    scriptBridge: {} as ExcelScriptBridge,
    uiBridge: {} as ExcelUiBridge,
    ...overrides,
  });
  return target;
}

describe("addExcelExecutors", () => {
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
          guidance: "WPS 正则提取使用 REGEXP；不要使用 Excel 方言的 REGEXEXTRACT/REGEXREPLACE/REGEXTEST 函数名",
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
    const workbookBridge = {
      writeRange: vi.fn(async () => undefined),
    } as unknown as ExcelWorkbookBridge;
    const target = createExcelExecutors({ workbookBridge });

    const result = await target.get("range.write")!.execute({
      sheetName: "Sheet1",
      range: "A1:B1",
      values: "[[1,2]]",
    });

    expect(result).toEqual({ success: true, data: "写入成功" });
    expect(workbookBridge.writeRange).toHaveBeenCalledWith("Sheet1", "A1:B1", [[1, 2]]);
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
      values: [['=LET(t,A1,—LEFT(t,2))']],
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
