/**
 * ExcelVbaComBridge — VBA COM 桥接实现
 *
 * 通过 PowerShell COM 自动化执行 VBA 宏和模块操作。
 */

import type { ExcelVbaBridge } from "../../contracts/excel";
import { executePowerShell, psVar } from "../../../automation/powershell";
import type { ExcelComBridge } from "./excelComBridge";

export class ExcelVbaComBridge implements ExcelVbaBridge {
  private comBridge: ExcelComBridge;

  constructor(comBridge: ExcelComBridge) {
    this.comBridge = comBridge;
  }

  /** 获取当前 COM ProgID（通过 comBridge 的 host 判断） */
  private getProgId(): string {
    return this.comBridge.host === "wps" ? "Ket.Application" : "Excel.Application";
  }

  async detectCapabilities(): Promise<{ supported: boolean; version?: string }> {
    try {
      const result = await this.executeVbaCheck();
      return { supported: result, version: result ? "VBA" : undefined };
    } catch {
      return { supported: false };
    }
  }

  async runMacro(macroName: string, args?: unknown[]): Promise<unknown> {
    try {
      const progId = this.getProgId();
      const argsStr = args ? args.map((a) => JSON.stringify(a)).join(",") : "";
      await executePowerShell(`
        ${psVar("_macroName", macroName)}
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("${progId}")
        $excel.Run($_macroName${argsStr ? `, ${argsStr}` : ""})
      `);
      return { success: true };
    } catch (err: any) {
      throw new Error(`运行宏失败: ${err.message}`);
    }
  }

  async writeModule(moduleName: string, code: string): Promise<void> {
    try {
      const progId = this.getProgId();
      // 将 VBA 代码编码为 Base64 传递，避免多行代码和特殊字符在 PowerShell 中转义问题
      const codeBase64 = Buffer.from(code, "utf16le").toString("base64");
      await executePowerShell(`
        ${psVar("_moduleName", moduleName)}
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("${progId}")
        $wb = $excel.ActiveWorkbook
        $vbProject = $wb.VBProject
        $module = $vbProject.VBComponents.Add(1)
        $module.Name = $_moduleName
        $codeBytes = [System.Convert]::FromBase64String("${codeBase64}")
        $codeText = [System.Text.Encoding]::Unicode.GetString($codeBytes)
        $module.CodeModule.AddFromString($codeText)
      `);
    } catch (err: any) {
      throw new Error(`写入模块失败: ${err.message}`);
    }
  }

  async executeCode(code: string): Promise<unknown> {
    try {
      const moduleName = `TempModule_${Date.now()}`;
      await this.writeModule(moduleName, code);
      await this.runMacro("Main");
      const progId = this.getProgId();
      await executePowerShell(`
        ${psVar("_moduleName", moduleName)}
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("${progId}")
        $wb = $excel.ActiveWorkbook
        $vbProject = $wb.VBProject
        $module = $vbProject.VBComponents.Item($_moduleName)
        $vbProject.VBComponents.Remove($module)
      `);
      return { success: true };
    } catch (err: any) {
      throw new Error(`执行代码失败: ${err.message}`);
    }
  }

  /** 检查 VBA 是否可用（供 ExcelScriptBridgeCom 调用） */
  async executeVbaCheck(): Promise<boolean> {
    try {
      const progId = this.getProgId();
      const result = await executePowerShell(`
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("${progId}")
        try {
          $vb = $excel.ActiveWorkbook.VBProject
          "VBA_AVAILABLE"
        } catch {
          "VBA_UNAVAILABLE"
        }
      `);
      return result.trim() === "VBA_AVAILABLE";
    } catch {
      return false;
    }
  }
}
