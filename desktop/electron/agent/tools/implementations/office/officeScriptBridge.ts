/**
 * OfficeScriptBridge — Word/PowerPoint 通用 PowerShell COM 脚本执行
 */

import type { OfficeScriptBridge as OfficeScriptBridgeInterface } from "../../contracts/office";
import { executePowerShell, psVar } from "../../../automation/powershell";
import { safeJsonParse } from "../../../automation/json";

const APP_PROG_IDS: Record<"word" | "presentation", string[]> = {
  word: ["Word.Application", "Kwps.Application", "Wps.Application"],
  presentation: ["PowerPoint.Application", "Wpp.Application", "Kwpp.Application"],
};

function progIdsLiteral(app: "word" | "presentation"): string {
  return "@(" + APP_PROG_IDS[app].map((id) => `'${id}'`).join(", ") + ")";
}

export class OfficeScriptBridge implements OfficeScriptBridgeInterface {
  async executeScript(app: "word" | "presentation", code: string): Promise<unknown> {
    try {
      if (!code.trim()) {
        throw new Error("脚本内容不能为空");
      }
      const result = await executePowerShell(`
${psVar("_code", code)}
$progIds = ${progIdsLiteral(app)}
$app = $null
$progId = $null
foreach ($id in $progIds) {
  try {
    $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject($id)
    $progId = $id
    break
  } catch {}
}
if ($null -eq $app) {
  throw '未找到已打开的 Office COM 应用，请先打开 Word 或 PowerPoint 文档'
}
$app.Visible = $true
$output = Invoke-Expression $_code | Out-String
[pscustomobject]@{
  success = $true
  app = '${app}'
  progId = $progId
  engine = 'powershell'
  output = $output.Trim()
} | ConvertTo-Json -Depth 5 -Compress
`, 90000);
      return safeJsonParse(result, "powershell", "执行 Office 脚本");
    } catch (err: any) {
      throw new Error(`执行 Office 脚本失败: ${err.message}`);
    }
  }
}
