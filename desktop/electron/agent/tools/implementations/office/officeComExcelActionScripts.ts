import { psVar } from "../../../automation/powershell";
import { parseOfficeLocator } from "../../officeCore/locator";
import type { OfficeActionInput } from "../../officeCore/types";
import { buildAcquireOfficeAppScript } from "./officeComPowerShell";
import {
  actionParamsScript,
  colorToOle,
  excelChartType,
  excelChangeDetail,
  excelChangeKind,
  outputPathForAction,
  psLiteralText,
  stringParam,
  validationFormulaParam,
} from "./officeComActionScriptHelpers";
import { buildChartAdvancedOperationScript } from "./officeComExcelChartScripts";
import { buildExcelFormulaGovernanceScript } from "./officeComExcelFormulaGovernanceScripts";
import { buildPowerQueryOperationScript } from "./officeComExcelPowerQueryScripts";
import { buildExcelPrintOperationScript } from "./officeComExcelPrintScripts";
import { buildWorkbookTemplateOperationScript } from "./officeComExcelTemplateScripts";
import { buildWorkbookObjectOperationScript } from "./officeComExcelWorkbookObjectScripts";

const EXCEL_PROG_IDS = ["Excel.Application", "Ket.Application"];

export function buildExcelScript(input: OfficeActionInput): string {
  const requestedHost = (stringParam(input, "host") || "").trim().toLowerCase();
  const progIds = requestedHost === "wps" || requestedHost === "ket"
    ? ["Ket.Application"]
    : requestedHost === "excel" || requestedHost === "microsoft" || requestedHost === "office"
      ? ["Excel.Application"]
      : EXCEL_PROG_IDS;
  const locator = parseOfficeLocator(input.target || "");
  const sheetName = locator.sheetName || stringParam(input, "sheetName") || "Sheet1";
  const rangeAddress = locator.address || stringParam(input, "range") || "A1";
  const targetExplicit = Boolean(locator.address || stringParam(input, "range"));
  const isExport = ["exportPdf", "exportSheetsToPdf", "snapshot"].includes(input.operation);
  const isReadOnly = [
    "traceFormulaDependencies",
    "inspectFormulaDependencies",
    "inspectFormulaBackups",
    "inspectFormulaProtection",
    "inspectPrintSettings",
    "inspectPowerQueries",
    "inspectCharts",
    "inspectWorkbookObjects",
    "captureWorkbookTemplate",
    "inspectWorkbookFormatting",
  ].includes(input.operation);
  const outputPath = outputPathForAction(
    input,
    input.operation === "snapshot"
      ? "snapshot.png"
      : input.operation === "exportPdf"
        ? "export.pdf"
        : input.operation === "exportSheetsToPdf"
          ? "sheets.pdf"
          : undefined,
  );

  return `
${psVar("_filePath", input.filePath!)}
${psVar("_outputPath", outputPath)}
${psVar("_sheetName", sheetName)}
${psVar("_rangeAddress", rangeAddress)}
${psVar("_validationFormula", validationFormulaParam(input))}
${psVar("_cfFormula", stringParam(input, "formula") || "TRUE")}
${psVar("_operation", input.operation)}
${actionParamsScript(input)}
$chartType = ${excelChartType(stringParam(input, "chartType"))}
$fillColor = ${colorToOle(stringParam(input, "fillColor") || "FFF2CC")}
$styleColor = ${colorToOle(stringParam(input, "headerColor") || "1F4E79")}
$app = $null
$workbook = $null
$sheet = $null
$range = $null
$openedWorkbook = $false
$createdApp = $false
$createdOfficeProcessId = [uint32]0
$ownsOfficeProcess = $false
$officeProcessIdsBefore = @()
$changes = @()
$operationData = [ordered]@{}
$targetExplicit = ${targetExplicit ? "$true" : "$false"}
try {
  $officeProcessIdsBefore = @(Get-Process -Name 'EXCEL', 'wps' -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
${buildAcquireOfficeAppScript({
    progIds,
    appKind: "excel",
    reuseAnyActive: false,
    missingMessage: "未找到可用的 Excel/WPS 表格 COM 应用",
    visible: false,
  })}
  if ($createdApp) {
    try {
      if (-not ('WenggeNativeWindow' -as [type])) {
        Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class WenggeNativeWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
      }
      [void][WenggeNativeWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd, [ref]$createdOfficeProcessId)
    } catch {}
    if ($createdOfficeProcessId -gt 0 -and [int]$createdOfficeProcessId -in $officeProcessIdsBefore) { $createdOfficeProcessId = [uint32]0 }
    if ($createdOfficeProcessId -eq 0) {
      try {
        $createdProcess = Get-Process -Name 'EXCEL', 'wps' -ErrorAction SilentlyContinue |
          Where-Object { [int]$_.Id -notin $officeProcessIdsBefore } |
          Sort-Object StartTime -Descending |
          Select-Object -First 1
        if ($null -ne $createdProcess) { $createdOfficeProcessId = [uint32]$createdProcess.Id }
      } catch {}
    }
    $ownsOfficeProcess = $createdOfficeProcessId -gt 0
    if ($ownsOfficeProcess -and $env:WENGGE_MANAGED_PROCESS_ID_FILE) {
      try { [IO.File]::WriteAllText($env:WENGGE_MANAGED_PROCESS_ID_FILE, [string]$createdOfficeProcessId) } catch {}
    }
  }
  $wantedPath = [IO.Path]::GetFullPath($_filePath)
  foreach ($candidate in $app.Workbooks) { try { if ([IO.Path]::GetFullPath([string]$candidate.FullName) -ieq $wantedPath) { $workbook = $candidate; break } } catch {} }
  if ($null -eq $workbook) { $workbook = $app.Workbooks.Open($_filePath); $openedWorkbook = $true }
  try { $sheet = $workbook.Worksheets.Item($_sheetName) } catch { throw "找不到工作表: $_sheetName" }
  $range = $sheet.Range($_rangeAddress)
  $null = . {
${excelOperationScript(input.operation)}
  }
  if (-not ${isReadOnly ? "$true" : "$false"} -and -not ${isExport ? "$true" : "$false"}) {
    if ($_outputPath -and $_outputPath -ne $_filePath) { $workbook.SaveAs($_outputPath) }
    else { $workbook.Save(); $_outputPath = $_filePath }
  }
  [pscustomobject]@{
    outputPath = $_outputPath
    changes = $changes
    data = $operationData
  } | ConvertTo-Json -Depth 10 -Compress
} finally {
  if ($null -ne $range) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($range) } catch {}; $range = $null }
  if ($null -ne $sheet) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet) } catch {}; $sheet = $null }
  if ($null -ne $workbook) {
    if ($openedWorkbook) { try { $workbook.Close($false) } catch {} }
    try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) } catch {}
    $workbook = $null
  }
  if ($null -ne $app) {
    if ($createdApp) {
      $workbooks = $null
      try {
        $workbooks = $app.Workbooks
        for ($index = $workbooks.Count; $index -ge 1; $index--) {
          $remainingWorkbook = $null
          try { $remainingWorkbook = $workbooks.Item($index); $remainingWorkbook.Close($false) } catch {}
          if ($null -ne $remainingWorkbook) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($remainingWorkbook) } catch {} }
        }
      } catch {}
      if ($null -ne $workbooks) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbooks) } catch {}; $workbooks = $null }
      try { $app.Quit() } catch {}
    }
    try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) } catch {}
    $app = $null
  }
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
  if ($ownsOfficeProcess) {
    try {
      $createdProcess = Get-Process -Id $createdOfficeProcessId -ErrorAction Stop
      if (-not $createdProcess.HasExited) { try { [void]$createdProcess.WaitForExit(3000) } catch {} }
      $createdProcess = Get-Process -Id $createdOfficeProcessId -ErrorAction SilentlyContinue
      if ($null -ne $createdProcess -and -not $createdProcess.HasExited) { Stop-Process -Id $createdOfficeProcessId -Force -ErrorAction SilentlyContinue }
    } catch {}
  }
  if ($env:WENGGE_MANAGED_PROCESS_ID_FILE) { try { Remove-Item -LiteralPath $env:WENGGE_MANAGED_PROCESS_ID_FILE -Force -ErrorAction SilentlyContinue } catch {} }
}
`;
}

function excelOperationScript(operation: string): string {
  const advanced = buildPowerQueryOperationScript(operation)
    || buildChartAdvancedOperationScript(operation)
    || buildWorkbookObjectOperationScript(operation)
    || buildWorkbookTemplateOperationScript(operation)
    || buildExcelPrintOperationScript(operation)
    || buildExcelFormulaGovernanceScript(operation);
  if (advanced) return advanced;

  const defaultChange = `
  $changes += [pscustomobject]@{ kind = '${psLiteralText(excelChangeKind(operation))}'; target = '${psLiteralText(operation)}'; detail = '${psLiteralText(excelChangeDetail(operation))}' }
`;
  switch (operation) {
    case "insertChart":
      return `
  $chartObjects = $sheet.ChartObjects()
  $chartObject = $chartObjects.Add($range.Left + $range.Width + 20, $range.Top, 420, 260)
  $chart = $chartObject.Chart
  $chart.SetSourceData($range)
  $chart.ChartType = $chartType
${defaultChange}`;
    case "setDataValidation":
      return `
  try { $range.Validation.Delete() } catch {}
  $range.Validation.Add(3, 1, 1, $_validationFormula)
  $range.Validation.IgnoreBlank = $true
  $range.Validation.InCellDropdown = $true
${defaultChange}`;
    case "applyConditionalFormatting":
      return `
  $condition = $range.FormatConditions.Add(2, 3, $_cfFormula)
  $condition.Interior.Color = $fillColor
${defaultChange}`;
    case "styleTable":
      return `
  $range.Font.Bold = $true
  $range.Borders.LineStyle = 1
  $range.Interior.Color = $styleColor
${defaultChange}`;
    case "createPivotTable":
      return `
  $pivotName = if ($actionParams.name) { [string]$actionParams.name } else { 'AI_Pivot_' + [DateTime]::Now.ToString('HHmmss') }
  $destination = if ($actionParams.destination) { [string]$actionParams.destination } else { $_sheetName + '!H3' }
  $parts = $destination -split '!', 2
  $destinationSheet = if ($parts.Count -eq 2) { $workbook.Worksheets.Item($parts[0].Trim("'")) } else { $sheet }
  $destinationAddress = if ($parts.Count -eq 2) { $parts[1] } else { $destination }
  $sourceAddress = $range.Address($true, $true, 1, $true)
  $cache = $workbook.PivotCaches().Create(1, $sourceAddress)
  $pivot = $cache.CreatePivotTable($destinationSheet.Range($destinationAddress), $pivotName)
  foreach ($fieldName in @($actionParams.rowFields | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })) { $pivot.PivotFields([string]$fieldName).Orientation = 1 }
  foreach ($fieldName in @($actionParams.columnFields | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })) { $pivot.PivotFields([string]$fieldName).Orientation = 2 }
  foreach ($fieldName in @($actionParams.filterFields | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })) { $pivot.PivotFields([string]$fieldName).Orientation = 3 }
  foreach ($field in @($actionParams.dataFields | Where-Object { $null -ne $_ })) {
    $sourceName = if ($field.name) { [string]$field.name } else { [string]$field }
    $caption = if ($field.caption) { [string]$field.caption } else { '汇总项: ' + $sourceName }
    $fn = switch ([string]$field.function) { 'average' { -4106 } 'count' { -4112 } 'max' { -4136 } 'min' { -4139 } default { -4157 } }
    [void]$pivot.AddDataField($pivot.PivotFields($sourceName), $caption, $fn)
  }
  $operationData.pivotName = $pivotName
  $operationData.destination = $destination
${defaultChange}`;
    case "refreshPivotTables":
      return `
  $refreshed = 0
  foreach ($ws in $workbook.Worksheets) {
    foreach ($pivot in $ws.PivotTables()) { $pivot.RefreshTable(); $refreshed++ }
  }
  if ($actionParams.refreshConnections -ne $false) { try { $workbook.RefreshAll() } catch {} }
  $operationData.refreshed = $refreshed
${defaultChange}`;
    case "addSlicer":
      return `
  if (-not $actionParams.field) { throw 'addSlicer 需要 params.field' }
  $pivotName = [string]$actionParams.pivotName
  $pivot = $null
  foreach ($ws in $workbook.Worksheets) {
    foreach ($candidate in $ws.PivotTables()) {
      if (-not $pivotName -or $candidate.Name -eq $pivotName) { $pivot = $candidate; break }
    }
    if ($null -ne $pivot) { break }
  }
  if ($null -eq $pivot) { throw '找不到可用于切片器的数据透视表' }
  $cache = try { $workbook.SlicerCaches.Add2($pivot, [string]$actionParams.field) } catch { $workbook.SlicerCaches.Add($pivot, [string]$actionParams.field) }
  $name = if ($actionParams.name) { [string]$actionParams.name } else { 'AI_Slicer_' + [DateTime]::Now.ToString('HHmmss') }
  $caption = if ($actionParams.caption) { [string]$actionParams.caption } else { [string]$actionParams.field }
  $left = if ($actionParams.left) { [double]$actionParams.left } else { $range.Left + $range.Width + 20 }
  $top = if ($actionParams.top) { [double]$actionParams.top } else { $range.Top }
  [void]$cache.Slicers.Add($sheet, $null, $name, $caption, $top, $left, 144, 180)
  $operationData.slicerName = $name
${defaultChange}`;
    case "exportPdf":
      return `
  $target = if ($actionParams.scope -eq 'sheet') { $sheet } else { $workbook }
  $target.ExportAsFixedFormat(0, $_outputPath)
  $changes += [pscustomobject]@{ kind = 'export'; target = $_outputPath; detail = '已导出 Excel PDF' }
`;
    case "snapshot":
      return `
  $range.CopyPicture(1, 2)
  $chartObject = $sheet.ChartObjects().Add($range.Left, $range.Top, [Math]::Max(1, $range.Width), [Math]::Max(1, $range.Height))
  try {
    $chartObject.Chart.Paste()
    [void]$chartObject.Chart.Export($_outputPath, 'PNG')
  } finally {
    $chartObject.Delete()
  }
  $changes += [pscustomobject]@{ kind = 'snapshot'; target = $_outputPath; detail = '已导出 Excel 区域快照' }
`;
    default:
      return "  throw \"不支持的 Excel COM 操作: $_operation\"";
  }
}
