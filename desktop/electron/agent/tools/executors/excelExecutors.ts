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
import { validateArgs } from "./validation";
import { toModelFacingSpreadsheetMetadata } from "./modelFacingMetadata";

export interface ExcelExecutorDeps {
  workbookBridge: ExcelWorkbookBridge;
  vbaBridge: ExcelVbaBridge;
  scriptBridge: ExcelScriptBridge;
  uiBridge: ExcelUiBridge;
}

const WPS_HOSTILE_FORMULA_CHARS: Record<string, string> = {
  "—": "长横线不是减号；请改成 ASCII 减号 '-'，双负号请写成 '--'",
  "–": "短横线不是减号；请改成 ASCII 减号 '-'，双负号请写成 '--'",
  "−": "数学负号不是减号；请改成 ASCII 减号 '-'，双负号请写成 '--'",
  "－": "全角减号不是减号；请改成 ASCII 减号 '-'，双负号请写成 '--'",
  "＋": "全角加号不是加号；请改成 ASCII 加号 '+'",
  "＊": "全角星号不是乘号；请改成 ASCII 星号 '*'",
  "／": "全角斜杠不是除号；请改成 ASCII 斜杠 '/'",
  "，": "全角逗号不能作为函数参数分隔符；请改成 ASCII 逗号 ','",
  "；": "全角分号不能作为函数参数分隔符；请改成 ASCII 分号 ';' 或当前环境支持的分隔符",
  "（": "全角左括号不能作为函数括号；请改成 ASCII 左括号 '('",
  "）": "全角右括号不能作为函数括号；请改成 ASCII 右括号 ')'",
  "：": "全角冒号不能作为区域引用符；请改成 ASCII 冒号 ':'",
  "！": "全角感叹号不能作为工作表引用符；请改成 ASCII 感叹号 '!'",
  "“": "弯引号不能作为公式字符串定界符；请改成 ASCII 双引号 '\"'",
  "”": "弯引号不能作为公式字符串定界符；请改成 ASCII 双引号 '\"'",
  "‘": "弯单引号不能作为工作表名定界符；请改成 ASCII 单引号 \"'\"",
  "’": "弯单引号不能作为工作表名定界符；请改成 ASCII 单引号 \"'\"",
  "\u200B": "公式中包含零宽空格，请删除该不可见字符",
  "\u200C": "公式中包含零宽非连接符，请删除该不可见字符",
  "\u200D": "公式中包含零宽连接符，请删除该不可见字符",
  "\uFEFF": "公式中包含 BOM/零宽不换行空格，请删除该不可见字符",
};

function validateFormulaTypographyForWps(values: unknown): string | null {
  for (const formula of collectFormulaStrings(values)) {
    const issue = findWpsHostileFormulaChar(formula);
    if (issue) {
      return [
        `WPS 公式解析风险：公式中包含 ${describeFormulaChar(issue.char)}。`,
        issue.message,
        "这类字符会触发 WPS 原生“公式有错误”弹窗，已在写入前拦截；请修正公式后重新调用 range.write。",
      ].join(" ");
    }
  }
  return null;
}

function collectFormulaStrings(value: unknown): string[] {
  if (typeof value === "string" && isFormulaLike(value)) return [value];
  if (!Array.isArray(value)) return [];
  return value.flatMap((item) => collectFormulaStrings(item));
}

function isFormulaLike(value: string): boolean {
  const trimmed = value.trimStart();
  return trimmed.startsWith("=") || trimmed.startsWith("+") || trimmed.startsWith("-");
}

function findWpsHostileFormulaChar(formula: string): { char: string; message: string } | null {
  let inDoubleQuotedText = false;
  let inSingleQuotedSheetName = false;
  for (let i = 0; i < formula.length; i++) {
    const char = formula[i];

    if (char === '"' && !inSingleQuotedSheetName) {
      if (inDoubleQuotedText && formula[i + 1] === '"') {
        i++;
        continue;
      }
      inDoubleQuotedText = !inDoubleQuotedText;
      continue;
    }

    if (char === "'" && !inDoubleQuotedText) {
      inSingleQuotedSheetName = !inSingleQuotedSheetName;
      continue;
    }

    if (inDoubleQuotedText || inSingleQuotedSheetName) continue;

    const message = WPS_HOSTILE_FORMULA_CHARS[char];
    if (message) return { char, message };
  }
  return null;
}

function describeFormulaChar(char: string): string {
  if (char.trim() === "") {
    return `不可见字符 U+${char.charCodeAt(0).toString(16).toUpperCase().padStart(4, "0")}`;
  }
  return `"${char}"`;
}

export function addExcelExecutors(target: Map<string, ToolExecutor>, deps: ExcelExecutorDeps): void {
  const { workbookBridge, vbaBridge, scriptBridge, uiBridge } = deps;

  target.set("workbook.inspect", {
    name: "workbook.inspect",
    execute: async (_args: Record<string, unknown>) => {
      const result = await workbookBridge.inspectWorkbook();
      return { success: true, data: toModelFacingSpreadsheetMetadata(result) };
    },
  });

  target.set("range.read", {
    name: "range.read",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { sheetName: "string", range: "string" });
      if (err) return { success: false, error: err };
      const requestedExpand = normalizeRangeReadExpand(args.expand);
      if (requestedExpand instanceof Error) return { success: false, error: requestedExpand.message };
      const shouldAutoDetectSpill =
        isOmittedExpand(args.expand) && isSingleCellRange(args.range as string);
      const effectiveExpand = shouldAutoDetectSpill ? "spill" : requestedExpand;
      const result = await workbookBridge.readRange(
        args.sheetName as string,
        args.range as string,
        effectiveExpand,
      );
      return {
        success: true,
        data: requestedExpand === "none" ? result.values : result,
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
      const formulaTypographyError = validateFormulaTypographyForWps(args.values);
      if (formulaTypographyError) {
        return { success: false, error: formulaTypographyError };
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

function isOmittedExpand(value: unknown): boolean {
  return value === undefined || value === null || value === "";
}

function isSingleCellRange(range: string): boolean {
  return /^\$?[A-Z]{1,3}\$?\d+$/i.test(range.trim());
}
