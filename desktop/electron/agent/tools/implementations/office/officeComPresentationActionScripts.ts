import { psVar } from "../../../automation/powershell";
import type { OfficeActionInput } from "../../officeCore/types";
import { buildAcquireOfficeAppScript } from "./officeComPowerShell";
import {
  actionParamsScript,
  colorToOle,
  outputPathForAction,
  pptChartType,
  presentationChangeDetail,
  presentationChangeKind,
  presentationDeleteSlideIndexesLiteral,
  psLiteralText,
  stringParam,
  targetSlideIndex,
} from "./officeComActionScriptHelpers";
import { buildPresentationBrandingOperationScript } from "./officeComPresentationBrandingScripts";
import { buildPresentationInspectionOperationScript } from "./officeComPresentationInspectionScripts";
import { buildPresentationPlaybackOperationScript } from "./officeComPresentationPlaybackScripts";
import { buildPresentationLinkedContentOperationScript } from "./officeComLinkedContentScripts";

const PRESENTATION_PROG_IDS = ["PowerPoint.Application", "Wpp.Application", "Kwpp.Application"];

export function buildPresentationScript(input: OfficeActionInput): string {
  const requestedHost = (stringParam(input, "host") || "").trim().toLowerCase();
  const progIds = requestedHost === "wps" || requestedHost === "wpp"
    ? ["Wpp.Application", "Kwpp.Application"]
    : requestedHost === "powerpoint" || requestedHost === "ppt"
      ? ["PowerPoint.Application"]
      : PRESENTATION_PROG_IDS;
  const isExport = input.operation === "snapshot" || input.operation === "exportHandouts";
  const isReadOnly = ["inspectPresentationTheme", "inspectSlideElements", "inspectAnimations", "inspectSpeakerNotes", "inspectLinkedOfficeContent"].includes(input.operation);
  const outputPath = outputPathForAction(input, input.operation === "snapshot" ? "snapshot.png" : input.operation === "exportHandouts" ? "handouts.pdf" : undefined);

  return `
${psVar("_filePath", input.filePath!)}
${psVar("_outputPath", outputPath)}
${psVar("_imagePath", stringParam(input, "imagePath") || "")}
${psVar("_operation", input.operation)}
${actionParamsScript(input)}
$slideIndex = ${targetSlideIndex(input.target)}
$deleteSlideIndexes = ${presentationDeleteSlideIndexesLiteral(input)}
$accentColor = ${colorToOle(stringParam(input, "accentColor") || "1F4E79")}
$backgroundColor = ${colorToOle(stringParam(input, "backgroundColor") || "FFFFFF")}
$chartType = ${pptChartType(stringParam(input, "chartType"))}
$app = $null
$pres = $null
$slide = $null
$openedPresentation = $false
$createdApp = $false
$createdOfficeProcessId = [uint32]0
$ownsOfficeProcess = $false
$pendingThemePackageUpdate = $false
$pendingNotesPackageUpdates = @()
$officeProcessIdsBefore = @()
$changes = @()
$operationData = [ordered]@{}
function Save-PresentationFile($presentation, [string]$path, [bool]$saveAs) {
  $lastError = $null
  foreach ($attempt in 1..4) {
    try {
      if ($saveAs) { $presentation.SaveAs($path) } else { $presentation.Save() }
      return
    } catch {
      $lastError = $_
      Start-Sleep -Milliseconds (250 * $attempt)
    }
  }
  throw $lastError
}
try {
  $officeProcessIdsBefore = @(Get-Process -Name 'POWERPNT', 'wpp', 'wps' -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
${buildAcquireOfficeAppScript({
    progIds,
    appKind: "presentation",
    reuseAnyActive: false,
    missingMessage: "未找到可用的 PowerPoint/WPS 演示 COM 应用",
    visible: -1,
  })}
  if ($createdApp) {
    try {
      if (-not ('WenggeNativeWindow' -as [type])) {
        Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class WenggeNativeWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
      }
      [void][WenggeNativeWindow]::GetWindowThreadProcessId([IntPtr]$app.HWND, [ref]$createdOfficeProcessId)
    } catch {}
    if ($createdOfficeProcessId -gt 0 -and [int]$createdOfficeProcessId -in $officeProcessIdsBefore) { $createdOfficeProcessId = [uint32]0 }
    if ($createdOfficeProcessId -eq 0) {
      try {
        $createdProcess = Get-Process -Name 'POWERPNT', 'wpp', 'wps' -ErrorAction SilentlyContinue |
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
  try { $app.DisplayAlerts = 1 } catch {}
  $operationData.progId = $progId
  $wantedPath = [IO.Path]::GetFullPath($_filePath)
  foreach ($candidate in $app.Presentations) { try { if ([IO.Path]::GetFullPath([string]$candidate.FullName) -ieq $wantedPath) { $pres = $candidate; break } } catch {} }
  if ($null -eq $pres) { $pres = $app.Presentations.Open($_filePath); $openedPresentation = $true }
  if ($slideIndex -gt $pres.Slides.Count) { throw "幻灯片序号超出范围: $slideIndex" }
  $slide = $pres.Slides.Item($slideIndex)
  $null = . {
${presentationOperationScript(input.operation)}
  }
  foreach ($operationObjectName in @('shape', 'picture', 'chart', 'table')) {
    $operationObject = Get-Variable -Name $operationObjectName -ValueOnly -ErrorAction SilentlyContinue
    if ($null -ne $operationObject -and [Runtime.InteropServices.Marshal]::IsComObject($operationObject)) {
      try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($operationObject) } catch {}
      Set-Variable -Name $operationObjectName -Value $null
    }
  }
  if (-not ${isExport ? "$true" : "$false"} -and -not ${isReadOnly ? "$true" : "$false"}) {
    if ($_outputPath -and $_outputPath -ne $_filePath) { Save-PresentationFile $pres $_outputPath $true }
    else { Save-PresentationFile $pres $_filePath $false; $_outputPath = $_filePath }
  }
  [pscustomobject]@{ outputPath = $_outputPath; changes = $changes; data = $operationData } |
    ConvertTo-Json -Depth 10 -Compress
} finally {
  if ($null -ne $slide) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($slide) } catch {}; $slide = $null }
  if ($null -ne $pres) {
    if ($openedPresentation) { try { $pres.Close() } catch {} }
    try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($pres) } catch {}
    $pres = $null
  }
  if ($null -ne $app) {
    if ($createdApp) {
      $presentations = $null
      try {
        $presentations = $app.Presentations
        for ($index = $presentations.Count; $index -ge 1; $index--) {
          $remainingPresentation = $null
          try { $remainingPresentation = $presentations.Item($index); $remainingPresentation.Close() } catch {}
          if ($null -ne $remainingPresentation) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($remainingPresentation) } catch {} }
        }
      } catch {}
      if ($null -ne $presentations) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($presentations) } catch {}; $presentations = $null }
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
  if ($pendingThemePackageUpdate) {
    $themePackagePath = if ($_outputPath -and $_outputPath -ne '') { $_outputPath } else { $_filePath }
    Update-PresentationThemePackage $themePackagePath @($actionParams.themeColors)
  }
  if ($pendingNotesPackageUpdates.Count -gt 0) {
    $notesPackagePath = if ($_outputPath -and $_outputPath -ne '') { $_outputPath } else { $_filePath }
    Update-PresentationNotesPackage $notesPackagePath @($pendingNotesPackageUpdates)
  }
  if ($env:WENGGE_MANAGED_PROCESS_ID_FILE) { try { Remove-Item -LiteralPath $env:WENGGE_MANAGED_PROCESS_ID_FILE -Force -ErrorAction SilentlyContinue } catch {} }
}
`;
}

function presentationOperationScript(operation: string): string {
  const advanced = buildPresentationInspectionOperationScript(operation)
    || buildPresentationBrandingOperationScript(operation)
    || buildPresentationPlaybackOperationScript(operation)
    || buildPresentationLinkedContentOperationScript(operation);
  if (advanced) return advanced;
  const defaultChange = `
  $changes += [pscustomobject]@{ kind = '${psLiteralText(presentationChangeKind(operation))}'; target = 'slide:' + $slideIndex; detail = '${psLiteralText(presentationChangeDetail(operation))}' }
`;
  switch (operation) {
    case "snapshot":
      return `
  $slide.Export($_outputPath, 'PNG')
  $changes += [pscustomobject]@{ kind = 'snapshot'; target = $_outputPath; detail = '已导出幻灯片快照' }
`;
    case "deleteSlides":
      return `
  if ($deleteSlideIndexes.Count -eq 0) { throw 'deleteSlides 需要 params.slides、params.from/to 或 target: slide:2-6' }
  foreach ($idx in ($deleteSlideIndexes | Sort-Object -Descending)) {
    if ($idx -lt 1 -or $idx -gt $pres.Slides.Count) { throw "幻灯片序号超出范围: $idx" }
    if ($pres.Slides.Count -le 1) { throw 'deleteSlides 至少需要保留一张幻灯片' }
    $pres.Slides.Item($idx).Delete()
  }
${defaultChange}`;
    case "insertChart":
      return `
  $shape = $slide.Shapes.AddChart2(201, $chartType, 80, 120, 520, 300)
  $shape.Name = if ($actionParams.name) { [string]$actionParams.name } else { 'AI Chart' }
${defaultChange}`;
    case "insertTable":
      return `
  $rows = if ($actionParams.rows) { [Math]::Max(1, [int]$actionParams.rows) } elseif ($actionParams.values) { @($actionParams.values).Count } else { 2 }
  $columns = if ($actionParams.columns) { [Math]::Max(1, [int]$actionParams.columns) } elseif ($actionParams.values -and @($actionParams.values).Count -gt 0) { @($actionParams.values[0]).Count } else { 2 }
  $left = if ($null -ne $actionParams.left) { [double]$actionParams.left } else { 80 }
  $top = if ($null -ne $actionParams.top) { [double]$actionParams.top } else { 120 }
  $width = if ($null -ne $actionParams.width) { [double]$actionParams.width } else { 520 }
  $height = if ($null -ne $actionParams.height) { [double]$actionParams.height } else { 220 }
  $shape = $slide.Shapes.AddTable($rows, $columns, $left, $top, $width, $height)
  $shape.Name = if ($actionParams.name) { [string]$actionParams.name } else { 'AI Table' }
  for ($rowIndex = 1; $rowIndex -le $rows; $rowIndex++) {
    for ($columnIndex = 1; $columnIndex -le $columns; $columnIndex++) {
      if ($actionParams.values -and $rowIndex -le @($actionParams.values).Count -and $columnIndex -le @($actionParams.values[$rowIndex - 1]).Count) {
        $shape.Table.Cell($rowIndex, $columnIndex).Shape.TextFrame.TextRange.Text = [string]$actionParams.values[$rowIndex - 1][$columnIndex - 1]
      }
    }
  }
${defaultChange}`;
    case "replacePictureSlot":
      return `
  if ([string]::IsNullOrWhiteSpace($_imagePath)) { throw 'replacePictureSlot 需要 params.imagePath' }
  if ($actionParams.shapeName) { try { $slide.Shapes.Item([string]$actionParams.shapeName).Delete() } catch {} }
  $left = if ($null -ne $actionParams.left) { [double]$actionParams.left } else { 80 }
  $top = if ($null -ne $actionParams.top) { [double]$actionParams.top } else { 120 }
  $width = if ($null -ne $actionParams.width) { [double]$actionParams.width } else { 520 }
  $height = if ($null -ne $actionParams.height) { [double]$actionParams.height } else { 300 }
  $picture = $slide.Shapes.AddPicture($_imagePath, $false, $true, $left, $top, $width, $height)
  if ($actionParams.name) { $picture.Name = [string]$actionParams.name }
  if ($actionParams.preserveAspectRatio -eq $true) { try { $picture.LockAspectRatio = -1 } catch {} }
${defaultChange}`;
    case "alignShapes":
      return `
  foreach ($shape in $slide.Shapes) {
    try { if ($shape.Left -lt 80) { $shape.Left = 80 }; if ($shape.Top -lt 80) { $shape.Top = 80 } } catch {}
  }
${defaultChange}`;
    case "normalizeLayouts":
      return `
  foreach ($shape in $slide.Shapes) {
    try { if ($shape.Width -gt 600) { $shape.Width = 600 }; if ($shape.Height -gt 360) { $shape.Height = 360 } } catch {}
  }
${defaultChange}`;
    case "applyTheme":
      return `
  foreach ($shape in $slide.Shapes) {
    try { if ($shape.HasTextFrame -and $shape.TextFrame.HasText) { $shape.TextFrame.TextRange.Font.Color.RGB = $accentColor } } catch {}
  }
${defaultChange}`;
    default:
      return "  throw \"不支持的 PowerPoint COM 操作: $_operation\"";
  }
}
