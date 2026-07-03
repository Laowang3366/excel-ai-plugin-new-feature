/**
 * Excel 工具执行器
 *
 * 只注册 Excel/WPS 工作簿、范围、公式、脚本和 UI 控件相关工具。
 */

import type { ToolExecutor } from "../../shared/types";
import type {
  ExcelWorkbookBridge,
  ExcelVbaBridge,
  ExcelScriptBridge,
  ExcelUiBridge,
  RangeReadExpandMode,
} from "../contracts/excel";
import { searchExcelFunctions } from "../data/excelFunctionCatalog";
import { validateArgs } from "./validation";

export interface ExcelExecutorDeps {
  workbookBridge: ExcelWorkbookBridge;
  vbaBridge: ExcelVbaBridge;
  scriptBridge: ExcelScriptBridge;
  uiBridge: ExcelUiBridge;
}

export function addExcelExecutors(target: Map<string, ToolExecutor>, deps: ExcelExecutorDeps): void {
  const { workbookBridge, vbaBridge, scriptBridge, uiBridge } = deps;

  target.set("workbook.inspect", {
    name: "workbook.inspect",
    execute: async (_args: Record<string, unknown>) => {
      const result = await workbookBridge.inspectWorkbook();
      return { success: true, data: result };
    },
  });

  target.set("range.read", {
    name: "range.read",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { sheetName: "string", range: "string" });
      if (err) return { success: false, error: err };
      const expand = normalizeRangeReadExpand(args.expand);
      if (expand instanceof Error) return { success: false, error: expand.message };
      const result = await workbookBridge.readRange(args.sheetName as string, args.range as string, expand);
      return {
        success: true,
        data: expand === "none" ? result.values : result,
      };
    },
  });

  target.set("range.write", {
    name: "range.write",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { sheetName: "string", range: "string" });
      if (err) return { success: false, error: err };
      if (args.values === undefined || args.values === null) {
        return { success: false, error: "缺少必填参数: values" };
      }
      if (typeof args.values === "string") {
        try {
          args.values = JSON.parse(args.values);
        } catch {
          return { success: false, error: "参数 values 应为数组，但收到了字符串且无法解析为 JSON" };
        }
      }
      if (!Array.isArray(args.values)) {
        return { success: false, error: `参数 values 应为数组，实际为 ${typeof args.values}` };
      }
      await workbookBridge.writeRange(args.sheetName as string, args.range as string, args.values as unknown[][]);
      return { success: true, data: "写入成功" };
    },
  });

  target.set("range.clear", {
    name: "range.clear",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { sheetName: "string", range: "string" });
      if (err) return { success: false, error: err };
      await workbookBridge.clearRange(args.sheetName as string, args.range as string);
      return { success: true, data: "清除成功" };
    },
  });

  target.set("selection.get", {
    name: "selection.get",
    execute: async (_args: Record<string, unknown>) => {
      const selection = await workbookBridge.getSelection();
      return { success: true, data: selection };
    },
  });

  target.set("formula.context", {
    name: "formula.context",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { sheetName: "string" });
      if (err) return { success: false, error: err };
      const context = await workbookBridge.getFormulaContext(args.sheetName as string, args.range as string | undefined);
      return { success: true, data: context };
    },
  });

  target.set("vba.runMacro", {
    name: "vba.runMacro",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { macroName: "string" });
      if (err) return { success: false, error: err };
      const result = await vbaBridge.runMacro(args.macroName as string, args.args as unknown[] | undefined);
      return { success: true, data: result };
    },
  });

  target.set("vba.writeModule", {
    name: "vba.writeModule",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { moduleName: "string", code: "string" });
      if (err) return { success: false, error: err };
      await vbaBridge.writeModule(args.moduleName as string, args.code as string);
      return { success: true, data: "模块写入成功" };
    },
  });

  target.set("formula.search", {
    name: "formula.search",
    execute: async (args: Record<string, unknown>) => {
      const query = ((args.query as string) || "").toLowerCase();
      const category = (args.category as string) || "";
      const results = searchExcelFunctions(query, category);
      return { success: true, data: { query: args.query as string, results } };
    },
  });

  target.set("sheet.operation", {
    name: "sheet.operation",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { operation: "string", sheetName: "string" });
      if (err) return { success: false, error: err };
      const result = await workbookBridge.sheetOperation(
        args.operation as string,
        args.sheetName as string,
        args as Record<string, unknown>
      );
      return { success: true, data: result };
    },
  });

  target.set("script.detect", {
    name: "script.detect",
    execute: async (_args: Record<string, unknown>) => {
      const env = await scriptBridge.detectEnvironment();
      return { success: true, data: env };
    },
  });

  target.set("script.execute", {
    name: "script.execute",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { code: "string" });
      if (err) return { success: false, error: err };
      const result = await scriptBridge.executeScript(
        args.code as string,
        args.language as string | undefined
      );
      return { success: true, data: result };
    },
  });

  target.set("ui.addControl", {
    name: "ui.addControl",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { sheetName: "string", controlType: "string", name: "string" });
      if (err) return { success: false, error: err };
      const result = await uiBridge.addControl({
        sheetName: args.sheetName as string,
        controlType: args.controlType as string,
        name: args.name as string,
        left: args.left as number,
        top: args.top as number,
        width: args.width as number,
        height: args.height as number,
        caption: args.caption as string | undefined,
        macroName: args.macroName as string | undefined,
        linkedCell: args.linkedCell as string | undefined,
      });
      return { success: true, data: result };
    },
  });

  target.set("ui.removeControl", {
    name: "ui.removeControl",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { sheetName: "string", name: "string" });
      if (err) return { success: false, error: err };
      await uiBridge.removeControl(args.sheetName as string, args.name as string);
      return { success: true, data: "控件已删除" };
    },
  });

  target.set("ui.listControls", {
    name: "ui.listControls",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { sheetName: "string" });
      if (err) return { success: false, error: err };
      const controls = await uiBridge.listControls(args.sheetName as string);
      return { success: true, data: controls };
    },
  });

  target.set("ui.createForm", {
    name: "ui.createForm",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { formName: "string", caption: "string" });
      if (err) return { success: false, error: err };
      const result = await uiBridge.createForm({
        formName: args.formName as string,
        caption: args.caption as string,
        width: args.width as number | undefined,
        height: args.height as number | undefined,
        controls: args.controls as Array<Record<string, unknown>> | undefined,
        eventCode: args.eventCode as string | undefined,
      });
      return { success: true, data: result };
    },
  });

  target.set("ui.addMenu", {
    name: "ui.addMenu",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { menuBar: "string", caption: "string", macroName: "string" });
      if (err) return { success: false, error: err };
      const result = await uiBridge.addMenu({
        menuBar: args.menuBar as string,
        caption: args.caption as string,
        macroName: args.macroName as string,
        beforeId: args.beforeId as number | undefined,
        faceId: args.faceId as number | undefined,
      });
      return { success: true, data: result };
    },
  });

  target.set("workbook.open", {
    name: "workbook.open",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { filePath: "string" });
      if (err) return { success: false, error: err };
      const result = await workbookBridge.openWorkbook(args.filePath as string);
      return { success: result.success, data: result };
    },
  });

  target.set("workbook.create", {
    name: "workbook.create",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { filePath: "string" });
      if (err) return { success: false, error: err };
      const result = await workbookBridge.createWorkbook(
        args.filePath as string,
        args.sheetNames as string[] | undefined
      );
      return { success: result.success, data: result };
    },
  });

  target.set("workbook.save", {
    name: "workbook.save",
    execute: async (args: Record<string, unknown>) => {
      const result = await workbookBridge.saveWorkbook(args.saveAsPath as string | undefined);
      return { success: result.success, data: result };
    },
  });

  target.set("workbook.switch", {
    name: "workbook.switch",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { workbookName: "string" });
      if (err) return { success: false, error: err };
      const result = await workbookBridge.switchWorkbook(args.workbookName as string);
      return { success: result.success, data: result };
    },
  });
}

function normalizeRangeReadExpand(value: unknown): RangeReadExpandMode | Error {
  if (value === undefined || value === null || value === "") return "none";
  if (typeof value !== "string") return new Error("参数 expand 必须是 none、spill、currentArray 或 currentRegion");
  if (value === "none" || value === "spill" || value === "currentArray" || value === "currentRegion") {
    return value;
  }
  return new Error("参数 expand 必须是 none、spill、currentArray 或 currentRegion");
}
