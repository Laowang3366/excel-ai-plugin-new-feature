/**
 * Excel 宏能力工具执行器
 *
 * 只注册工作簿内部 VBA/WPS JSA 宏相关工具。
 */

import type { ToolExecutor } from "../../shared/types";
import type {
  ExcelWorkbookBridge,
  ExcelVbaBridge,
  WpsJsaBridge,
  WorkbookMacroLanguage,
} from "../contracts/excel";
import { validateArgs } from "./validation";

export interface ExcelMacroExecutorDeps {
  workbookBridge: ExcelWorkbookBridge;
  vbaBridge: ExcelVbaBridge;
  jsaBridge: WpsJsaBridge;
}

export function addExcelMacroExecutors(
  target: Map<string, ToolExecutor>,
  deps: ExcelMacroExecutorDeps,
): void {
  const { workbookBridge, vbaBridge, jsaBridge } = deps;

  target.set("macro.detect", {
    name: "macro.detect",
    execute: async (_args: Record<string, unknown>) => {
      const [hostInfo, vba, javascript] = await Promise.all([
        workbookBridge.getHostInfo(),
        vbaBridge.detectCapabilities(),
        jsaBridge.detectCapabilities(),
      ]);
      const capabilities = [
        {
          language: "vba" as const,
          supported: vba.supported,
          ready: vba.supported,
          internal: true as const,
          engine: "VBA" as const,
          reason: vba.reason,
        },
        javascript,
      ];
      const available = capabilities.filter((item) => item.supported);
      const ready = available.filter((item) => item.ready);
      return {
        success: true,
        data: {
          host: hostInfo.host,
          recommended: ready[0]?.language ?? available[0]?.language ?? "none",
          available,
          unavailable: capabilities.filter((item) => !item.supported),
        },
      };
    },
  });

  target.set("macro.run", {
    name: "macro.run",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { language: "string", macroName: "string" });
      if (err) return { success: false, error: err };
      if (args.language !== "vba") {
        return {
          success: false,
          error: "macro.run 当前仅支持 vba；WPS JSA 只提供写入和回读校验",
        };
      }
      const result = await vbaBridge.runMacro(
        args.macroName as string,
        args.args as unknown[] | undefined,
      );
      return { success: true, data: result };
    },
  });

  target.set("macro.write", {
    name: "macro.write",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, {
        language: "string",
        code: "string",
        entryPoint: "string",
      });
      if (err) return { success: false, error: err };
      const language = normalizeMacroLanguage(args.language);
      if (language instanceof Error) return { success: false, error: language.message };

      if (language === "vba") {
        const vbaErr = validateArgs(args, { moduleName: "string" });
        if (vbaErr) return { success: false, error: vbaErr };
        const result = await vbaBridge.writeModule(args.moduleName as string, args.code as string, {
          entryPoint: args.entryPoint as string,
          save: true,
          saveAsPath: args.saveAsPath as string | undefined,
        });
        return { success: true, data: { language, ...result } };
      }

      if (args.saveAsPath !== undefined) {
        return {
          success: false,
          error: "saveAsPath 仅支持 VBA；WPS JSA 保存到当前工作簿",
        };
      }
      const result = await jsaBridge.writeCode(args.code as string, {
        entryPoint: args.entryPoint as string,
        save: true,
      });
      return { success: true, data: result };
    },
  });
}

function normalizeMacroLanguage(value: unknown): WorkbookMacroLanguage | Error {
  if (value === "vba" || value === "javascript") return value;
  return new Error("参数 language 必须是 vba 或 javascript");
}
