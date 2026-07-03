/**
 * Excel/WPS 区域和选区能力。
 *
 * 关联模块：
 * - excelComBridge.ts: 对外保留 range/selection 公共方法。
 * - automation: 提供脚本执行、变量注入和 JSON 解析。
 * - rangeValueUtils.ts: 提供写入区域前的二维数组规范化。
 */

import { safeJsonParse } from "../../../automation/json";
import { jsVar } from "../../../automation/jscript";
import {
  executePowerShell,
  psVar,
} from "../../../automation/powershell";
import { pyVar } from "../../../automation/python";
import { executeSmart } from "../../../automation/scriptEngine";
import type { SpreadsheetHost } from "./connectionMetadata";
import { normalize2D } from "./rangeValueUtils";
import type { RangeReadExpandMode, RangeReadResult } from "../../contracts/excel";

export interface RangeOperationDeps {
  ensureConnected: (retries?: number) => Promise<SpreadsheetHost | null>;
  getProgId: () => string;
}

export async function readRangeOperation(
  deps: RangeOperationDeps,
  sheetName: string,
  range: string,
  expand: RangeReadExpandMode = "none"
): Promise<RangeReadResult> {
  const host = await deps.ensureConnected(0);
  if (!host) throw new Error("未连接到 Excel/WPS，请先在侧边栏点击连接");

  try {
    const progId = deps.getProgId();
    const psScript = `
${psVar("_sheetName", sheetName)}
${psVar("_range", range)}
${psVar("_expand", expand)}
$excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
$wb = $excel.ActiveWorkbook
$ws = $wb.Sheets.Item($_sheetName)
$range = $ws.Range($_range)
$readRange = $range
$expanded = $false
try {
  $mode = ([string]$_expand).ToLowerInvariant()
  if ($mode -eq 'spill') {
    try {
      $spillRange = $range.Cells.Item(1, 1).SpillingToRange
      if ($null -ne $spillRange) {
        $readRange = $spillRange
        $expanded = $true
      }
    } catch {
      try {
        $arrayRange = $range.Cells.Item(1, 1).CurrentArray
        if ($null -ne $arrayRange -and ([string]$arrayRange.Address(0, 0, 1, 0)) -ne ([string]$range.Address(0, 0, 1, 0))) {
          $readRange = $arrayRange
          $expanded = $true
        }
      } catch {}
    }
  } elseif ($mode -eq 'currentarray') {
    try {
      $arrayRange = $range.Cells.Item(1, 1).CurrentArray
      if ($null -ne $arrayRange) {
        $readRange = $arrayRange
        $expanded = (([string]$arrayRange.Address(0, 0, 1, 0)) -ne ([string]$range.Address(0, 0, 1, 0)))
      }
    } catch {}
  } elseif ($mode -eq 'currentregion') {
    try {
      $regionRange = $range.CurrentRegion
      if ($null -ne $regionRange) {
        $readRange = $regionRange
        $expanded = (([string]$regionRange.Address(0, 0, 1, 0)) -ne ([string]$range.Address(0, 0, 1, 0)))
      }
    } catch {}
  }
} catch {}
${rangeValue2RowsPowerShell()}
$rows = Convert-RangeValue2ToRows $readRange.Value2
[pscustomobject]@{
  values = $rows
  address = [string]$readRange.Address(0, 0, 1, 0)
  expanded = [bool]$expanded
  expandMode = [string]$_expand
} | ConvertTo-Json -Depth 8 -Compress
`;

    const result = await executePowerShell(psScript, 90000);
    const parsed = safeJsonParse<RangeReadResult | unknown[][]>(result, "powershell", "读取范围");
    if (Array.isArray(parsed)) {
      return { values: normalizeReadRows(parsed), address: range, expanded: false, expandMode: expand };
    }
    return {
      values: normalizeReadRows(parsed.values),
      address: parsed.address,
      expanded: parsed.expanded,
      expandMode: parsed.expandMode ?? expand,
    };
  } catch (err: any) {
    throw new Error(`读取范围失败: ${err.message}`);
  }
}

function normalizeReadRows(values: unknown): unknown[][] {
  if (Array.isArray(values)) {
    if (values.length > 0 && !Array.isArray(values[0])) {
      return [values];
    }
    return values as unknown[][];
  }
  return [[values]];
}

function rangeValue2RowsPowerShell(): string {
  return `
function Convert-CellValue($value) {
  if ($null -eq $value) { return "" }
  return $value
}
function Convert-RangeValue2ToRows($values) {
  $rows = [System.Collections.Generic.List[object]]::new()
  if ($values -is [System.Array]) {
    if ($values.Rank -eq 2) {
      for ($r = $values.GetLowerBound(0); $r -le $values.GetUpperBound(0); $r++) {
        $row = [System.Collections.Generic.List[object]]::new()
        for ($c = $values.GetLowerBound(1); $c -le $values.GetUpperBound(1); $c++) {
          [void]$row.Add((Convert-CellValue ($values.GetValue($r, $c))))
        }
        [void]$rows.Add($row.ToArray())
      }
      return $rows.ToArray()
    }
    for ($i = $values.GetLowerBound(0); $i -le $values.GetUpperBound(0); $i++) {
      [void]$rows.Add(@((Convert-CellValue ($values.GetValue($i)))))
    }
    return $rows.ToArray()
  }
  [void]$rows.Add(@((Convert-CellValue $values)))
  return $rows.ToArray()
}
`;
}

export async function writeRangeOperation(
  deps: RangeOperationDeps,
  sheetName: string,
  range: string,
  values: unknown[][]
): Promise<void> {
  const host = await deps.ensureConnected();
  if (!host) throw new Error("未连接到 Excel/WPS，请先在侧边栏点击连接");

  const normalized = normalize2D(values);

  try {
    const progId = deps.getProgId();
    const valuesJson = JSON.stringify(normalized);
    const valuesB64 = Buffer.from(valuesJson, "utf-8").toString("base64");

    const pythonScript = `
${pyVar("_sheetName", sheetName)}
${pyVar("_range", range)}
import json as _json
_values = _json.loads(base64.b64decode("${valuesB64}").decode("utf-8"))
s = wb.sheets[_sheetName]
rng = s.range(_range)
for r, row in enumerate(_values):
    for c, val in enumerate(row):
        rng.cells.item(r, c).value = val
print("OK")
`;

    const flatB64: string[] = [];
    for (const row of normalized) {
      for (const cell of row) {
        const cellStr = cell === null || cell === undefined ? "" : String(cell);
        flatB64.push(Buffer.from(cellStr, "utf16le").toString("base64"));
      }
    }
    let jsAssign = "";
    let idx = 0;
    for (let r = 0; r < normalized.length; r++) {
      for (let c = 0; c < (normalized[r]?.length || 0); c++) {
        jsAssign += `rng.Cells.Item(${r + 1}, ${c + 1}) = (function() { var n = new ActiveXObject("MSXML2.DOMDocument").createElement("b64"); n.DataType = "bin.base64"; n.Text = "${flatB64[idx]}"; var b = n.NodeTypedValue; var s = new ActiveXObject("ADODB.Stream"); s.Type = 1; s.Open(); s.Write(b); s.Position = 0; s.Type = 2; s.Charset = "Unicode"; var v = s.ReadText(); s.Close(); return v; })();\n`;
        idx++;
      }
    }
    const jscriptScript = `
${jsVar("_sheetName", sheetName)}
${jsVar("_range", range)}
var excel = GetObject("", "${progId}");
var wb = excel.ActiveWorkbook;
var ws = wb.Sheets.Item(_sheetName);
var rng = ws.Range(_range);
${jsAssign}
WScript.Echo("OK");
`;

    let psAssign = "";
    idx = 0;
    for (let r = 0; r < normalized.length; r++) {
      for (let c = 0; c < (normalized[r]?.length || 0); c++) {
        psAssign += `$startRange.Cells.Item(${r + 1}, ${c + 1}) = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String('${flatB64[idx]}'))\n`;
        idx++;
      }
    }
    const psScript = `
${psVar("_sheetName", sheetName)}
${psVar("_range", range)}
$ErrorActionPreference = 'Stop'
try {
  $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
  $wb = $excel.ActiveWorkbook
  $ws = $wb.Sheets.Item($_sheetName)
  $startRange = $ws.Range($_range)
${psAssign}
} catch {
  Write-Error $_.Exception.Message
  exit 1
}
`;

    await executeSmart(pythonScript, jscriptScript, psScript, 90000, { preferPython: false });
  } catch (err: any) {
    throw new Error(`写入范围失败: ${err.message}`);
  }
}

export async function clearRangeOperation(
  deps: RangeOperationDeps,
  sheetName: string,
  range: string
): Promise<void> {
  const host = await deps.ensureConnected();
  if (!host) throw new Error("未连接到 Excel/WPS，请先在侧边栏点击连接");

  try {
    const progId = deps.getProgId();
    const pythonScript = `
${pyVar("_sheetName", sheetName)}
${pyVar("_range", range)}
s = wb.sheets[_sheetName]
s.range(_range).clear()
print("OK")
`;

    const jscriptScript = `
${jsVar("_sheetName", sheetName)}
${jsVar("_range", range)}
var excel = GetObject("", "${progId}");
var wb = excel.ActiveWorkbook;
var ws = wb.Sheets.Item(_sheetName);
ws.Range(_range).Clear();
WScript.Echo("OK");
`;

    const psScript = `
${psVar("_sheetName", sheetName)}
${psVar("_range", range)}
$excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
$wb = $excel.ActiveWorkbook
$ws = $wb.Sheets.Item($_sheetName)
$ws.Range($_range).Clear()
`;

    await executeSmart(pythonScript, jscriptScript, psScript, 90000, { preferPython: false });
  } catch (err: any) {
    throw new Error(`清除范围失败: ${err.message}`);
  }
}

export async function getSelectionOperation(
  deps: RangeOperationDeps,
  _readRange: (sheetName: string, range: string) => Promise<unknown[][]>
): Promise<{ address: string; values: unknown[][]; sheetName: string }> {
  const host = await deps.ensureConnected(0);
  if (!host) throw new Error("未连接到 Excel/WPS，请先在侧边栏点击连接");

  try {
    const progId = deps.getProgId();
    const psScript = `
$excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
$sel = $excel.Selection
if ($null -eq $sel) { throw '当前没有可用选区' }
$address = $sel.Address(0, 0, 1, 0)
$sheetName = $sel.Worksheet.Name
${rangeValue2RowsPowerShell()}
$rows = Convert-RangeValue2ToRows $sel.Value2
[pscustomobject]@{ address = [string]$address; sheetName = [string]$sheetName; values = $rows } | ConvertTo-Json -Depth 8 -Compress
`;

    const result = await executePowerShell(psScript, 90000);
    const parsed = safeJsonParse<{ address: string; values: unknown[][]; sheetName: string }>(result, "powershell", "获取选区");
    if (Array.isArray(parsed.values) && parsed.values.length > 0 && !Array.isArray(parsed.values[0])) {
      parsed.values = [parsed.values];
    }
    return parsed;
  } catch (err: any) {
    throw new Error(`获取选区失败: ${err.message}`);
  }
}

export async function getSelectionAddressOperation(
  deps: RangeOperationDeps
): Promise<{ address: string; sheetName: string }> {
  const host = await deps.ensureConnected(0);
  if (!host) throw new Error("未连接到 Excel/WPS，请先在侧边栏点击连接");

  try {
    const progId = deps.getProgId();
    const result = await executePowerShell(`
$excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
$sel = $excel.Selection
if ($null -eq $sel) { throw '当前没有可用选区' }
$address = [string]$sel.Address(0, 0, 1, 0)
$sheetName = [string]$sel.Worksheet.Name
[pscustomobject]@{ address = $address; sheetName = $sheetName } | ConvertTo-Json -Compress
`, 15000);

    return safeJsonParse<{ address: string; sheetName: string }>(result, "powershell", "获取选区地址");
  } catch (err: any) {
    throw new Error(`获取选区地址失败: ${err.message}`);
  }
}
