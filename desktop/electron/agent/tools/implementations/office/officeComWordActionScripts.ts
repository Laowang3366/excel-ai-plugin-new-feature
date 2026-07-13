import { psVar } from "../../../automation/powershell";
import type { OfficeActionInput } from "../../officeCore/types";
import { buildAcquireOfficeAppScript } from "./officeComPowerShell";
import {
  actionParamsScript,
  normalizeHeadingLevel,
  outputPathForAction,
  psLiteralText,
  stringParam,
  wordChangeDetail,
  wordChangeKind,
} from "./officeComActionScriptHelpers";
import { buildWordContentControlOperationScript } from "./officeComWordContentControlScripts";
import { buildWordFormattingOperationScript } from "./officeComWordFormattingScripts";
import { buildWordMailMergeOperationScript } from "./officeComWordMailMergeScripts";
import { buildWordReferenceOperationScript } from "./officeComWordReferenceScripts";
import { buildWordRevisionOperationScript } from "./officeComWordRevisionScripts";
import { buildWordLinkedContentOperationScript } from "./officeComLinkedContentScripts";

const WORD_PROG_IDS = ["Word.Application", "Kwps.Application", "Wps.Application"];

export function buildWordScript(input: OfficeActionInput): string {
  const requestedHost = (stringParam(input, "host") || "").trim().toLowerCase();
  const progIds = requestedHost === "wps" || requestedHost === "kwps"
    ? ["Kwps.Application", "Wps.Application"]
    : requestedHost === "word" || requestedHost === "microsoft" || requestedHost === "office"
      ? ["Word.Application"]
      : WORD_PROG_IDS;
  const isExport = ["snapshot", "exportPdf", "mailMerge", "batchMailMerge", "compareDocuments"].includes(input.operation);
  const isReadOnly = ["inspectDocumentFormatting", "inspectReferences", "inspectRevisions", "inspectContentControls", "inspectLinkedOfficeContent"].includes(input.operation);
  const outputPath = outputPathForAction(
    input,
    input.operation === "snapshot"
      ? "preview.pdf"
      : input.operation === "exportPdf"
        ? "export.pdf"
        : input.operation === "compareDocuments"
          ? "comparison.docx"
          : input.operation === "mailMerge" || input.operation === "batchMailMerge"
          ? "merged.docx"
          : undefined,
  );
  const headingLevel = normalizeHeadingLevel(input.params?.level);

  return `
${psVar("_filePath", input.filePath!)}
${psVar("_outputPath", outputPath)}
${psVar("_startsWith", stringParam(input, "startsWith") || "")}
${psVar("_headerFooterText", stringParam(input, "text") || "")}
${psVar("_headerFooterKind", stringParam(input, "kind") === "footer" ? "footer" : "header")}
${psVar("_imagePath", stringParam(input, "imagePath") || "")}
${psVar("_operation", input.operation)}
${actionParamsScript(input)}
$headingLevel = ${headingLevel}
$app = $null
$doc = $null
$openedDocument = $false
$createdApp = $false
$createdOfficeProcessId = [uint32]0
$ownsOfficeProcess = $false
$officeProcessIdsBefore = @()
$changes = @()
$operationData = [ordered]@{}
try {
  $officeProcessIdsBefore = @(Get-Process -Name 'WINWORD', 'wps' -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
${buildAcquireOfficeAppScript({
    progIds,
    appKind: "word",
    reuseAnyActive: false,
    missingMessage: "未找到可用的 Word/WPS 文字 COM 应用",
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
        $createdProcess = Get-Process -Name 'WINWORD', 'wps' -ErrorAction SilentlyContinue |
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
  try { $app.DisplayAlerts = 0 } catch {}
  try { $app.Options.UpdateLinksAtOpen = $false } catch {}
  $wantedPath = [IO.Path]::GetFullPath($_filePath)
  foreach ($candidate in $app.Documents) { try { if ([IO.Path]::GetFullPath([string]$candidate.FullName) -ieq $wantedPath) { $doc = $candidate; break } } catch {} }
  if ($null -eq $doc) { $doc = $app.Documents.Open($_filePath, $false, $false, $false); $openedDocument = $true }
  $null = . {
${wordOperationScript(input.operation)}
  }
  if (-not ${isExport ? "$true" : "$false"} -and -not ${isReadOnly ? "$true" : "$false"}) {
    if ($_outputPath -and $_outputPath -ne $_filePath) { $doc.SaveAs2($_outputPath) }
    else { $doc.Save(); $_outputPath = $_filePath }
  }
  [pscustomobject]@{ outputPath = $_outputPath; changes = $changes; data = $operationData } |
    ConvertTo-Json -Depth 10 -Compress
} finally {
  if ($null -ne $doc) {
    if ($openedDocument) { try { $doc.Close(0) } catch {} }
    try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc) } catch {}
    $doc = $null
  }
  if ($null -ne $app) {
    if ($createdApp) {
      $documents = $null
      try {
        $documents = $app.Documents
        for ($index = $documents.Count; $index -ge 1; $index--) {
          $remainingDocument = $null
          try { $remainingDocument = $documents.Item($index); $remainingDocument.Close(0) } catch {}
          if ($null -ne $remainingDocument) { try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($remainingDocument) } catch {} }
        }
      } catch {}
      if ($null -ne $documents) { try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($documents) } catch {}; $documents = $null }
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
      if (-not $createdProcess.HasExited) {
        try { Wait-Process -Id $createdOfficeProcessId -Timeout 3 -ErrorAction Stop } catch {}
      }
      $createdProcess = Get-Process -Id $createdOfficeProcessId -ErrorAction SilentlyContinue
      if ($null -ne $createdProcess -and -not $createdProcess.HasExited) { Stop-Process -Id $createdOfficeProcessId -Force -ErrorAction SilentlyContinue }
    } catch {}
  }
  if ($env:WENGGE_MANAGED_PROCESS_ID_FILE) { try { Remove-Item -LiteralPath $env:WENGGE_MANAGED_PROCESS_ID_FILE -Force -ErrorAction SilentlyContinue } catch {} }
}
`;
}

function wordOperationScript(operation: string): string {
  const advanced = buildWordFormattingOperationScript(operation)
    || buildWordReferenceOperationScript(operation)
    || buildWordRevisionOperationScript(operation)
    || buildWordMailMergeOperationScript(operation)
    || buildWordContentControlOperationScript(operation)
    || buildWordLinkedContentOperationScript(operation);
  if (advanced) return advanced;

  const defaultChange = `
  $changes += [pscustomobject]@{ kind = '${psLiteralText(wordChangeKind(operation))}'; target = '${psLiteralText(operation)}'; detail = '${psLiteralText(wordChangeDetail(operation))}' }
`;
  switch (operation) {
    case "styleTables":
      return `
  foreach ($table in $doc.Tables) {
    $table.Borders.Enable = 1
    if ($table.Rows.Count -gt 0) {
      $table.Rows.Item(1).Range.Bold = $true
      $table.Rows.Item(1).Shading.BackgroundPatternColor = 15773696
    }
    try { $table.AutoFitBehavior(1) } catch {}
  }
${defaultChange}`;
    case "setHeaderFooter":
      return `
  foreach ($section in $doc.Sections) {
    if ($_headerFooterKind -eq 'footer') { $section.Footers.Item(1).Range.Text = $_headerFooterText }
    else { $section.Headers.Item(1).Range.Text = $_headerFooterText }
  }
${defaultChange}`;
    case "insertOrReplaceImage":
      return `
  if ([string]::IsNullOrWhiteSpace($_imagePath)) { throw 'insertOrReplaceImage 需要 params.imagePath' }
  $insertRange = $doc.Range([Math]::Max(0, $doc.Content.End - 1), [Math]::Max(0, $doc.Content.End - 1))
  [void]$doc.InlineShapes.AddPicture($_imagePath, $false, $true, $insertRange)
${defaultChange}`;
    case "snapshot":
    case "exportPdf":
      return `
  $doc.ExportAsFixedFormat($_outputPath, 17)
  $changes += [pscustomobject]@{ kind = 'export'; target = $_outputPath; detail = '已导出 Word PDF' }
`;
    default:
      return "  throw \"不支持的 Word COM 操作: $_operation\"";
  }
}
