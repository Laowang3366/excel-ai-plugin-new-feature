/**
 * Excel/WPS 公式上下文能力。
 *
 * 关联模块：
 * - excelComBridge.ts: 对外保留 getFormulaContext 门面。
 * - prompts/sections/formulaAssistantPrompt.ts: 公式助手会依赖该工具理解已有公式。
 */

import { safeJsonParse } from "../../../automation/json";
import { executePowerShell, psVar } from "../../../automation/powershell";
import type { SpreadsheetHost } from "./connectionMetadata";

export interface FormulaOperationDeps {
  ensureConnected: (retries?: number) => Promise<SpreadsheetHost | null>;
  getProgId: () => string;
}

export async function getFormulaContextOperation(
  deps: FormulaOperationDeps,
  sheetName: string,
  range?: string
): Promise<unknown> {
  const host = await deps.ensureConnected(0);
  if (!host) throw new Error("未连接到 Excel/WPS，请先在侧边栏点击连接");

  try {
    const progId = deps.getProgId();
    const rangeArg = range || "A1:Z100";

    const psScript = `
${psVar("_sheetName", sheetName)}
${psVar("_range", rangeArg)}
function Convert-ColumnNumberToName([int]$columnNumber) {
  $name = ""
  while ($columnNumber -gt 0) {
    $columnNumber--
    $name = [char](65 + ($columnNumber % 26)) + $name
    $columnNumber = [math]::Floor($columnNumber / 26)
  }
  return $name
}
function Get-MatrixValue($values, [int]$rowOffset, [int]$colOffset) {
  if ($values -is [System.Array]) {
    if ($values.Rank -eq 2) {
      return $values.GetValue(
        $values.GetLowerBound(0) + $rowOffset,
        $values.GetLowerBound(1) + $colOffset
      )
    }
    return $values.GetValue($values.GetLowerBound(0) + $rowOffset)
  }
  return $values
}
$excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
$wb = $excel.ActiveWorkbook
$ws = $wb.Sheets.Item($_sheetName)
$range = $ws.Range($_range)
try {
  $formulaValues = $range.Formula2
} catch {
  $formulaValues = $range.Formula
}
$cellValues = $range.Value2
$startRow = [int]$range.Row
$startCol = [int]$range.Column
$rowCount = [int]$range.Rows.Count
$colCount = [int]$range.Columns.Count
$formulas = [System.Collections.Generic.List[object]]::new()
for ($r = 0; $r -lt $rowCount; $r++) {
  for ($c = 0; $c -lt $colCount; $c++) {
    $formula = Get-MatrixValue $formulaValues $r $c
    if ($formula -and ([string]$formula).StartsWith("=")) {
      $rowNumber = $startRow + $r
      $columnName = Convert-ColumnNumberToName ($startCol + $c)
      $value = Get-MatrixValue $cellValues $r $c
      [void]$formulas.Add([pscustomobject]@{
        cell = "$columnName$rowNumber"
        formula = [string]$formula
        value = if ($null -eq $value) { "" } else { [string]$value }
      })
    }
  }
}
ConvertTo-Json -InputObject $formulas.ToArray() -Depth 5 -Compress
`;

    const result = await executePowerShell(psScript, 90000);
    const parsed = safeJsonParse<unknown[]>(result, "powershell", "获取公式上下文");
    if (!Array.isArray(parsed)) return [parsed];
    return parsed;
  } catch (err: any) {
    throw new Error(`获取公式上下文失败: ${err.message}`);
  }
}
