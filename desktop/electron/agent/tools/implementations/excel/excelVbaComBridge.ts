/**
 * ExcelVbaComBridge — VBA COM 桥接实现
 *
 * 通过 PowerShell COM 自动化执行 VBA 宏和模块操作。
 */

import type {
  ExcelVbaBridge,
  VbaModuleWriteOptions,
  VbaModuleWriteResult,
} from "../../contracts/excel";
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

  async detectCapabilities(): Promise<{
    supported: boolean;
    version?: string;
    host?: "excel" | "wps";
    reason?: string;
  }> {
    try {
      const result = await this.executeVbaCheck();
      return {
        supported: result,
        version: result ? "VBA" : undefined,
        host: this.comBridge.host === "wps" ? "wps" : "excel",
        reason: result ? undefined : "无法访问活动工作簿的 VBA 工程，请确认已安装 VBA 并允许访问 VBA 工程对象模型",
      };
    } catch (err: any) {
      return {
        supported: false,
        host: this.comBridge.host === "wps" ? "wps" : "excel",
        reason: err.message,
      };
    }
  }

  async runMacro(macroName: string, args?: unknown[]): Promise<unknown> {
    if (!macroName.trim()) {
      throw new Error("宏名称不能为空");
    }
    if ((args?.length ?? 0) > 30) {
      throw new Error("宏参数不能超过 30 个");
    }
    try {
      const progId = this.getProgId();
      const argBindings = (args ?? []).map((arg, index) => [
        psVar(`_argJson${index}`, JSON.stringify(arg) ?? "null"),
        `$_arg${index} = ConvertFrom-Json -InputObject $_argJson${index}`,
      ].join("\n")).join("\n");
      const argList = (args ?? []).map((_, index) => `$_arg${index}`).join(", ");
      const result = await executePowerShell(`
        ${psVar("_macroName", macroName)}
        ${argBindings}
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("${progId}")
        $returnValue = $excel.Run($_macroName${argList ? `, ${argList}` : ""})
        [PSCustomObject]@{
          invoked = $true
          macroName = $_macroName
          returnValue = $returnValue
        } | ConvertTo-Json -Depth 5 -Compress
      `);
      return JSON.parse(result);
    } catch (err: any) {
      throw new Error(`运行宏失败: ${err.message}`);
    }
  }

  async writeModule(
    moduleName: string,
    code: string,
    options: VbaModuleWriteOptions = {}
  ): Promise<VbaModuleWriteResult> {
    const normalizedModuleName = moduleName.trim();
    const normalizedCode = normalizeVbaSource(code);
    const entryPoint = options.entryPoint?.trim();

    if (!normalizedModuleName || normalizedModuleName.length > 31) {
      throw new Error("模块名称必须为 1 到 31 个字符");
    }
    if (!normalizedCode) {
      throw new Error("VBA 代码不能为空");
    }
    if (entryPoint && !hasVbaProcedure(normalizedCode, entryPoint)) {
      throw new Error(`VBA 代码中找不到入口过程: ${entryPoint}`);
    }

    try {
      const progId = this.getProgId();
      const shouldSave = options.save === true;
      const result = await executePowerShell(`
        ${psVar("_moduleName", normalizedModuleName)}
        ${psVar("_codeText", normalizedCode)}
        ${psVar("_requestedSaveAsPath", options.saveAsPath?.trim() ?? "")}
        $shouldSave = $${shouldSave ? "true" : "false"}

        function Normalize-Code([string]$text) {
          if ($null -eq $text) { return "" }
          $normalized = $text.Replace(([string][char]13 + [char]10), [string][char]10)
          return $normalized.Replace([string][char]13, [string][char]10).TrimEnd([char]10)
        }

        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("${progId}")
        $wb = $excel.ActiveWorkbook
        if ($null -eq $wb) { throw "当前没有活动工作簿" }

        try {
          $vbProject = $wb.VBProject
          $null = $vbProject.VBComponents.Count
        } catch {
          throw "无法访问 VBA 工程。请在 Office/WPS 信任中心开启‘信任对 VBA 工程对象模型的访问’，并确认已安装 VBA 组件"
        }

        $saveTarget = ""
        $saveFormat = 0
        if ($shouldSave) {
          if ($_requestedSaveAsPath) {
            $extension = [System.IO.Path]::GetExtension($_requestedSaveAsPath).ToLowerInvariant()
            if ($extension -notin @(".xlsm", ".xlsb", ".xls")) {
              throw "宏工作簿保存路径必须使用 .xlsm、.xlsb 或 .xls 扩展名"
            }
            $saveTarget = $_requestedSaveAsPath
            $saveFormat = if ($extension -eq ".xlsm") { 52 } elseif ($extension -eq ".xlsb") { 50 } else { 56 }
          } else {
            $currentExtension = [System.IO.Path]::GetExtension($wb.Name).ToLowerInvariant()
            if ($currentExtension -notin @(".xlsm", ".xlsb", ".xls")) {
              if (-not $wb.Path) {
                throw "当前工作簿尚未保存且不是宏格式，请通过 saveAsPath 指定 .xlsm 保存路径"
              }
              $baseName = [System.IO.Path]::GetFileNameWithoutExtension($wb.Name)
              $saveTarget = [System.IO.Path]::Combine($wb.Path, "$($baseName)-macro.xlsm")
              $suffix = 2
              while ([System.IO.File]::Exists($saveTarget)) {
                $saveTarget = [System.IO.Path]::Combine($wb.Path, "$($baseName)-macro-$suffix.xlsm")
                $suffix++
              }
              $saveFormat = 52
            }
          }
        }

        $module = $null
        foreach ($component in $vbProject.VBComponents) {
          if ($component.Name -eq $_moduleName) {
            $module = $component
            break
          }
        }

        $created = $false
        $oldCode = ""
        $mutationStarted = $false
        try {
          if ($null -eq $module) {
            $module = $vbProject.VBComponents.Add(1)
            $created = $true
            $mutationStarted = $true
            $module.Name = $_moduleName
          } elseif ($module.Type -ne 1) {
            throw "同名组件不是标准模块，不能覆盖: $_moduleName"
          } else {
            $oldLineCount = $module.CodeModule.CountOfLines
            if ($oldLineCount -gt 0) {
              $oldCode = $module.CodeModule.Lines(1, $oldLineCount)
              $module.CodeModule.DeleteLines(1, $oldLineCount)
            }
          }

          $mutationStarted = $true
          $module.CodeModule.AddFromString($_codeText)
          $writtenLineCount = $module.CodeModule.CountOfLines
          $writtenCode = if ($writtenLineCount -gt 0) { $module.CodeModule.Lines(1, $writtenLineCount) } else { "" }
          if ((Normalize-Code $writtenCode) -cne (Normalize-Code $_codeText)) {
            throw "VBA 模块源码回读不一致"
          }

          $probeName = "RunProbe"
          $probeModuleName = "WenggeProbe_" + [Guid]::NewGuid().ToString("N").Substring(0, 12)
          $probeModule = $vbProject.VBComponents.Add(1)
          try {
            $probeModule.Name = $probeModuleName
            $probeCode = "Public Sub $probeName()" + [Environment]::NewLine + "End Sub"
            $probeModule.CodeModule.AddFromString($probeCode)
            $escapedWorkbookName = $wb.Name.Replace("'", "''")
            $probeMacro = "'$escapedWorkbookName'!$probeModuleName.$probeName"
            $null = $excel.Run($probeMacro)
          } finally {
            if ($null -ne $probeModule) { $vbProject.VBComponents.Remove($probeModule) }
          }

          $finalLineCount = $module.CodeModule.CountOfLines
          $finalCode = if ($finalLineCount -gt 0) { $module.CodeModule.Lines(1, $finalLineCount) } else { "" }
          if ((Normalize-Code $finalCode) -cne (Normalize-Code $_codeText)) {
            throw "编译校验后 VBA 模块源码发生变化"
          }

          if ($shouldSave) {
            if ($saveTarget) {
              $oldDisplayAlerts = $excel.DisplayAlerts
              try {
                $excel.DisplayAlerts = $false
                $wb.SaveAs($saveTarget, $saveFormat)
              } finally {
                $excel.DisplayAlerts = $oldDisplayAlerts
              }
            } else {
              $wb.Save()
            }
          }

          [PSCustomObject]@{
            moduleName = $_moduleName
            created = $created
            lineCount = $finalLineCount
            sourceBase64 = [System.Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($finalCode))
            saved = $shouldSave
            workbookName = $wb.Name
            workbookPath = $wb.FullName
          } | ConvertTo-Json -Compress
        } catch {
          if ($mutationStarted) {
            try {
              if ($created -and $null -ne $module) {
                $vbProject.VBComponents.Remove($module)
              } elseif ($null -ne $module) {
                $count = $module.CodeModule.CountOfLines
                if ($count -gt 0) { $module.CodeModule.DeleteLines(1, $count) }
                if ($oldCode) { $module.CodeModule.AddFromString($oldCode) }
              }
            } catch {}
          }
          throw
        }
      `);

      const parsed = JSON.parse(result) as {
        moduleName: string;
        created: boolean;
        lineCount: number;
        sourceBase64: string;
        saved: boolean;
        workbookName: string;
        workbookPath: string;
      };
      const readback = Buffer.from(parsed.sourceBase64, "base64").toString("utf16le");
      if (normalizeVbaSource(readback) !== normalizedCode) {
        throw new Error("VBA 模块源码回读不一致");
      }

      return {
        moduleName: parsed.moduleName,
        created: parsed.created,
        lineCount: parsed.lineCount,
        sourceVerified: true,
        compileVerified: true,
        entryPoint,
        entryPointVerified: entryPoint ? hasVbaProcedure(readback, entryPoint) : true,
        saved: parsed.saved,
        workbookName: parsed.workbookName,
        workbookPath: parsed.workbookPath,
        host: this.comBridge.host === "wps" ? "wps" : "excel",
      };
    } catch (err: any) {
      throw new Error(`写入模块失败: ${err.message}`);
    }
  }

  /** 检查当前工作簿的内部 VBA 工程是否可访问 */
  async executeVbaCheck(): Promise<boolean> {
    try {
      const progId = this.getProgId();
      const result = await executePowerShell(`
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("${progId}")
        try {
          $vb = $excel.ActiveWorkbook.VBProject
          $null = $vb.VBComponents.Count
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

function normalizeVbaSource(code: string): string {
  return code.replace(/\r\n?/g, "\n").replace(/\n+$/g, "").trimStart();
}

function hasVbaProcedure(code: string, entryPoint: string): boolean {
  const procedureName = entryPoint.split(".").pop()?.trim();
  if (!procedureName) return false;
  const escapedName = procedureName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return new RegExp(
    `^\\s*(?:Public\\s+)?(?:Static\\s+)?(?:Sub|Function)\\s+${escapedName}\\b`,
    "im"
  ).test(code);
}
