import { psVar } from "../../../automation/powershell";
import { parseOfficeLocator } from "../../officeCore/locator";
import type { OfficeActionInput } from "../../officeCore/types";
import { actionParamsScript, defaultOutputPath, stringParam } from "./officeComActionScriptHelpers";
import { officeInstanceDiscoveryScript } from "./officeDocumentComBridge";
import { progIdsLiteral } from "./officeComPowerShell";

const CROSS_OFFICE_OPERATIONS = new Set([
  "exportRangeToWord",
  "exportRangeToPresentation",
  "buildReportPackage",
]);

export function isCrossOfficeOperation(operation: string): boolean {
  return CROSS_OFFICE_OPERATIONS.has(operation);
}

export function buildCrossOfficeScript(input: OfficeActionInput): string {
  validateIncrementalLinkIds(input);
  const sourceHost = (stringParam(input, "sourceHost") || stringParam(input, "host") || "").trim().toLowerCase();
  const wordHost = (stringParam(input, "wordHost") || "").trim().toLowerCase();
  const presentationHost = (stringParam(input, "presentationHost") || "").trim().toLowerCase();
  const excelProgIds = sourceHost === "wps" || sourceHost === "ket"
    ? ["Ket.Application"]
    : sourceHost === "excel" || sourceHost === "microsoft" || sourceHost === "office"
      ? ["Excel.Application"]
      : ["Excel.Application", "Ket.Application"];
  const wordProgIds = wordHost === "wps" || wordHost === "kwps"
    ? ["Kwps.Application", "Wps.Application"]
    : wordHost === "word" || wordHost === "microsoft" || wordHost === "office"
      ? ["Word.Application"]
      : ["Word.Application", "Kwps.Application", "Wps.Application"];
  const presentationProgIds = presentationHost === "wps" || presentationHost === "wpp"
    ? ["Wpp.Application", "Kwpp.Application"]
    : presentationHost === "powerpoint" || presentationHost === "ppt" || presentationHost === "microsoft" || presentationHost === "office"
      ? ["PowerPoint.Application"]
      : ["PowerPoint.Application", "Wpp.Application", "Kwpp.Application"];
  const locator = parseOfficeLocator(input.target || "");
  const sheetName = locator.sheetName || stringParam(input, "sheetName") || "Sheet1";
  const rangeAddress = locator.address || stringParam(input, "range") || "A1";
  const fallbackOutput = input.operation === "exportRangeToWord"
    ? defaultOutputPath(input.filePath!, "range.docx")
    : defaultOutputPath(input.filePath!, "range.pptx");
  const outputPath = input.outputPath || fallbackOutput;

  return `
${psVar("_filePath", input.filePath!)}
${psVar("_outputPath", outputPath)}
${psVar("_sheetName", sheetName)}
${psVar("_rangeAddress", rangeAddress)}
${psVar("_operation", input.operation)}
${actionParamsScript(input)}
${officeInstanceDiscoveryScript()}
$excel = $null
$workbook = $null
$openedWorkbook = $false
$excelCreatedApp = $false
$word = $null
$wordDoc = $null
$openedWordDocument = $false
$createdWordDocument = $false
$wordCreatedApp = $false
$powerPoint = $null
  $presentation = $null
$openedPresentationDocument = $false
$createdPresentationDocument = $false
$presentationCreatedApp = $false
  $excelOwnedProcessId = [uint32]0
  $wordOwnedProcessId = [uint32]0
  $presentationOwnedProcessId = [uint32]0
  $tempOutputs = @()
$publishedOutputs = @()
$replacedOutputs = @()
  $changes = @()
  $operationData = [ordered]@{}
  $linkedObjects = @()
function New-OfficeComObject([string[]]$progIds) {
  $lastError = $null
  foreach ($progId in $progIds) {
    foreach ($attempt in 1..3) {
      $candidate = $null
      try {
        $candidate = New-Object -ComObject $progId
        $null = [string]$candidate.Version
        Write-Output -NoEnumerate $candidate
        return
      } catch {
        $lastError = $_
        if ($null -ne $candidate) { try { $candidate.Quit() } catch {}; try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($candidate) } catch {} }
      }
      Start-Sleep -Milliseconds (200 * $attempt)
    }
  }
  $detail = if ($null -ne $lastError) { ': ' + [string]$lastError.Exception.Message } else { '' }
  throw '未找到所需的 Office/WPS COM 应用' + $detail
}
function Find-OfficeDocumentHandle([string]$appKind, [string]$filePath, [string]$instanceId) {
  if (-not $filePath) { return $null }
  $wantedPath = [IO.Path]::GetFullPath($filePath)
  $candidates = @(Get-AllOfficeDocumentHandles $appKind | Where-Object {
    $pathMatches = try { [IO.Path]::GetFullPath([string]$_.document.FullName) -ieq $wantedPath } catch { $false }
    $instanceMatches = -not $instanceId -or [string]$_.instanceId -eq $instanceId
    $pathMatches -and $instanceMatches
  })
  if ($candidates.Count -gt 1) { throw '找到多个 Office 文档候选，请传对应的 instanceId' }
  return $(if ($candidates.Count -eq 1) { $candidates[0] } else { $null })
}
function Get-OwnedOfficeProcessId($application, [string[]]$processNames, [int[]]$beforeIds) {
  $processId = [uint32]0
  try {
    if (-not ('WenggeCrossOfficeNativeWindow' -as [type])) {
      Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class WenggeCrossOfficeNativeWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
    }
    [void][WenggeCrossOfficeNativeWindow]::GetWindowThreadProcessId([IntPtr]$application.Hwnd, [ref]$processId)
  } catch {}
  if ($processId -gt 0 -and [int]$processId -in $beforeIds) { return [uint32]0 }
  if ($processId -eq 0) {
    foreach ($attempt in 1..5) {
      try {
        $created = Get-Process -Name $processNames -ErrorAction SilentlyContinue | Where-Object { [int]$_.Id -notin $beforeIds } | Sort-Object StartTime -Descending | Select-Object -First 1
        if ($null -ne $created) { $processId = [uint32]$created.Id; break }
      } catch {}
      Start-Sleep -Milliseconds 100
    }
  }
  return $processId
}
function Stop-OwnedOfficeProcess([uint32]$processId) {
  if ($processId -eq 0) { return }
  try {
    $process = Get-Process -Id $processId -ErrorAction Stop
    if (-not $process.HasExited) { try { [void]$process.WaitForExit(3000) } catch {} }
    $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
    if ($null -ne $process -and -not $process.HasExited) { Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue }
  } catch {}
}
function Register-OwnedOfficeProcess([uint32]$processId) {
  if ($processId -eq 0 -or -not $env:WENGGE_MANAGED_PROCESS_ID_FILE) { return }
  $ids = @()
  try { $ids = @(([string][IO.File]::ReadAllText($env:WENGGE_MANAGED_PROCESS_ID_FILE) | ConvertFrom-Json)) } catch {}
  $ids = @($ids + [int]$processId | ForEach-Object { [int]$_ } | Where-Object { $_ -gt 0 } | Select-Object -Unique)
  try { [IO.File]::WriteAllText($env:WENGGE_MANAGED_PROCESS_ID_FILE, ($ids | ConvertTo-Json -Compress)) } catch {}
}
function New-AtomicOutputPath([string]$destination) {
  $directory = [System.IO.Path]::GetDirectoryName($destination)
  if (-not [System.IO.Directory]::Exists($directory)) { [void][System.IO.Directory]::CreateDirectory($directory) }
  $name = [System.IO.Path]::GetFileNameWithoutExtension($destination)
  $extension = [System.IO.Path]::GetExtension($destination)
  return [System.IO.Path]::Combine($directory, "." + $name + "." + [Guid]::NewGuid().ToString('N') + $extension)
}
function Publish-AtomicOutput([string]$temporary, [string]$destination, [bool]$overwrite) {
  if ([System.IO.File]::Exists($destination)) {
    if (-not $overwrite) { throw "输出文件已存在，请设置 params.overwrite=true: $destination" }
    $previous = New-AtomicOutputPath $destination
    Move-Item -LiteralPath $destination -Destination $previous
    $script:replacedOutputs += [pscustomobject]@{ destination = $destination; backup = $previous }
  }
  Move-Item -LiteralPath $temporary -Destination $destination
  $script:publishedOutputs += $destination
}
function Copy-ExcelContent($sourceSheet, [string]$rangeAddress, [string]$sourceType, [string]$sourceName, [bool]$asPicture) {
  if ($sourceType -eq 'chart') {
    if (-not $sourceName) { throw '图表联动需要 params.chartName 或 section.chartName' }
    $chartObject = $sourceSheet.ChartObjects($sourceName)
    $sourceSheet.Activate()
    $chartObject.Activate()
    if ($asPicture) { $chartObject.Chart.CopyPicture() } else { $chartObject.Chart.ChartArea.Copy() }
  } else {
    $contentRange = $sourceSheet.Range($rangeAddress)
    if ($asPicture) { $contentRange.CopyPicture(1, 2) } else { $contentRange.Copy() }
  }
}
function Invoke-OfficeClipboardPaste([scriptblock]$copyAction, [scriptblock]$pasteAction) {
  for ($attempt = 1; $attempt -le 5; $attempt++) {
    & $copyAction
    Start-Sleep -Milliseconds 200
    try {
      $pasteResult = & $pasteAction
      return ,$pasteResult
    } catch {
      if ($_.Exception.HResult -ne -2147418111 -or $attempt -eq 5) { throw }
      Start-Sleep -Milliseconds 300
    }
  }
}
function Get-LinkKey([string]$linkId) {
  $safe = $linkId -replace '[^A-Za-z0-9_]', '_'
  if ($safe.Length -gt 16) { $safe = $safe.Substring(0, 16) }
  $sha = [Security.Cryptography.SHA256]::Create()
  try { $hash = [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($linkId))).Replace('-', '').Substring(0, 10) } finally { $sha.Dispose() }
  return $safe + '_' + $hash
}
function Set-ObjectTag($item, [string]$name, [string]$value) {
  try { $item.Tags.Delete($name) } catch {}
  try { $item.Tags.Add($name, $value) } catch {}
}
function Get-WordManagedIds($document) {
  try { return @(([string]$document.Variables.Item('WENGGE_MANAGED_LINK_IDS').Value | ConvertFrom-Json)) } catch { return @() }
}
function Set-WordManagedIds($document, [string[]]$ids) {
  $uniqueIds = @($ids | Select-Object -Unique)
  $json = ConvertTo-Json -InputObject ([object[]]$uniqueIds) -Compress
  try { $document.Variables.Item('WENGGE_MANAGED_LINK_IDS').Value = $json } catch { [void]$document.Variables.Add('WENGGE_MANAGED_LINK_IDS', $json) }
  try { $document.Variables.Item('WENGGE_MANIFEST_VERSION').Value = '1' } catch { [void]$document.Variables.Add('WENGGE_MANIFEST_VERSION', '1') }
}
function Get-WordLinkMetadataVariableName([string]$bookmarkName) { return 'WENGGE_META_' + $bookmarkName }
function Set-WordLinkMetadata($document, [string]$bookmarkName, [string]$metadata) {
  $variableName = Get-WordLinkMetadataVariableName $bookmarkName
  try { $document.Variables.Item($variableName).Value = $metadata } catch { [void]$document.Variables.Add($variableName, $metadata) }
}
function Remove-WordLinkMetadata($document, [string]$bookmarkName) {
  try { $document.Variables.Item((Get-WordLinkMetadataVariableName $bookmarkName)).Delete() } catch {}
}
function Find-WordInlineShapeAt($document, [int]$start) {
  $match = $null; $matchStart = [int]::MaxValue
  foreach ($candidate in $document.InlineShapes) {
    $candidateStart = try { [int]$candidate.Range.Start } catch { continue }
    if ($candidateStart -ge $start -and $candidateStart -lt $matchStart) { $match = $candidate; $matchStart = $candidateStart }
  }
  return $match
}
function Find-WordTableAt($document, [int]$start) {
  $match = $null; $matchStart = [int]::MaxValue
  foreach ($candidate in $document.Tables) {
    $candidateStart = try { [int]$candidate.Range.Start } catch { continue }
    if ($candidateStart -ge $start -and $candidateStart -lt $matchStart) { $match = $candidate; $matchStart = $candidateStart }
  }
  return $match
}
function Get-PresentationManagedIds($targetPresentation) {
  try { return @(([string]$targetPresentation.Tags.Item('WENGGE_MANAGED_LINK_IDS') | ConvertFrom-Json)) } catch { return @() }
}
function Set-PresentationManagedIds($targetPresentation, [string[]]$ids) {
  $uniqueIds = @($ids | Select-Object -Unique)
  Set-ObjectTag $targetPresentation 'WENGGE_MANAGED_LINK_IDS' (ConvertTo-Json -InputObject ([object[]]$uniqueIds) -Compress)
  Set-ObjectTag $targetPresentation 'WENGGE_MANIFEST_VERSION' '1'
}
function Find-WordManagedRange($document, [string]$linkId, [string]$targetBookmark, [string]$targetControlTag) {
  $bookmarkName = 'WgLink_' + (Get-LinkKey $linkId)
  if ($document.Bookmarks.Exists($bookmarkName)) { return [pscustomobject]@{ range = $document.Bookmarks.Item($bookmarkName).Range.Duplicate; bookmark = $bookmarkName; replace = $true } }
  if ($targetBookmark -and $document.Bookmarks.Exists($targetBookmark)) { return [pscustomobject]@{ range = $document.Bookmarks.Item($targetBookmark).Range.Duplicate; bookmark = $bookmarkName; replace = $true; templateBookmark = $targetBookmark } }
  if ($targetControlTag) {
    foreach ($control in $document.ContentControls) {
      if ([string]$control.Tag -ieq $targetControlTag -or [string]$control.Title -ieq $targetControlTag) {
        $range = $control.Range.Duplicate; if ($range.End -gt $range.Start) { $range.MoveEnd(1, -1) }
        return [pscustomobject]@{ range = $range; bookmark = $bookmarkName; replace = $true }
      }
    }
  }
  $end = [Math]::Max(0, $document.Content.End - 1)
  return [pscustomobject]@{ range = $document.Range($end, $end); bookmark = $bookmarkName; replace = $false }
}
function Add-ExcelContentToWord($document, $sourceSheet, [string]$rangeAddress, [string]$sourceType, [string]$sourceName, [string]$title, [bool]$linked, [string]$linkId, [string]$targetBookmark = '', [string]$targetControlTag = '') {
  $destination = Find-WordManagedRange $document $linkId $targetBookmark $targetControlTag
  $insert = $destination.range
  $start = [int]$insert.Start
  $inlineShapesBefore = [int]$document.InlineShapes.Count
  $deletedInlineShapes = 0
  if ([bool]$destination.replace) {
    $replaceEnd = [int]$insert.End
    for ($shapeIndex = $document.InlineShapes.Count; $shapeIndex -ge 1; $shapeIndex--) {
      $candidateShape = $document.InlineShapes.Item($shapeIndex)
      try { if ([int]$candidateShape.Range.Start -ge $start -and [int]$candidateShape.Range.Start -le $replaceEnd) { $candidateShape.Delete(); $deletedInlineShapes++ } } catch {}
    }
    for ($shapeIndex = $document.Shapes.Count; $shapeIndex -ge 1; $shapeIndex--) {
      $candidateShape = $document.Shapes.Item($shapeIndex)
      try { if ([int]$candidateShape.Anchor.Start -ge $start -and [int]$candidateShape.Anchor.Start -le $replaceEnd) { $candidateShape.Delete() } } catch {}
    }
    [void]$insert.Delete(); $insert.SetRange($start, $start)
  }
  if ($title) { $insert.InsertAfter($title + [Environment]::NewLine); $insert.Collapse(0) }
  $contentStart = [int]$insert.Start
  $contentEnd = $contentStart
  $linkedItem = $null
  $metadata = [pscustomobject]@{ version = 1; linkId = $linkId; source = $_filePath; sourceType = $sourceType; sourceName = $sourceName; range = $rangeAddress; managed = $true } | ConvertTo-Json -Compress
  if ($linked) {
    $beforeCount = [int]$document.InlineShapes.Count
    $null = Invoke-OfficeClipboardPaste { Copy-ExcelContent $sourceSheet $rangeAddress $sourceType $sourceName $false } { $insert.PasteSpecial(0, $true, 0, $false, 0) }
    if ($document.InlineShapes.Count -le $beforeCount) { throw 'Word 未创建 Excel 链接对象' }
    $linkedItem = Find-WordInlineShapeAt $document $contentStart
    if ($null -eq $linkedItem) { throw '无法定位刚插入的 Word Excel 链接对象' }
    try { $linkedItem.AlternativeText = 'WENGGE_MANIFEST:' + $metadata } catch {}
    $contentEnd = [int]$linkedItem.Range.End
  } elseif ($sourceType -eq 'chart') {
    $beforeCount = [int]$document.InlineShapes.Count
    $null = Invoke-OfficeClipboardPaste { Copy-ExcelContent $sourceSheet $rangeAddress $sourceType $sourceName $true } { $insert.PasteSpecial(0, $false, 0, $false, 9) }
    if ($document.InlineShapes.Count -gt $beforeCount) { $insertedShape = Find-WordInlineShapeAt $document $contentStart; if ($null -ne $insertedShape) { $contentEnd = [int]$insertedShape.Range.End } }
  } else {
    $beforeCount = [int]$document.Tables.Count
    $null = Invoke-OfficeClipboardPaste { Copy-ExcelContent $sourceSheet $rangeAddress $sourceType $sourceName $false } { $insert.PasteExcelTable($false, $false, $false) }
    if ($document.Tables.Count -gt $beforeCount) { $insertedTable = Find-WordTableAt $document $contentStart; if ($null -ne $insertedTable) { $contentEnd = [int]$insertedTable.Range.End } }
  }
  if ($contentEnd -lt $contentStart) { $contentEnd = $contentStart }
  $managedRange = $document.Range($start, $contentEnd)
  try { if ($document.Bookmarks.Exists([string]$destination.bookmark)) { $document.Bookmarks.Item([string]$destination.bookmark).Delete() }; [void]$document.Bookmarks.Add([string]$destination.bookmark, $managedRange) } catch {}
  Set-WordLinkMetadata $document ([string]$destination.bookmark) $metadata
  if ($destination.templateBookmark) { try { if ($document.Bookmarks.Exists([string]$destination.templateBookmark)) { $document.Bookmarks.Item([string]$destination.templateBookmark).Delete() }; [void]$document.Bookmarks.Add([string]$destination.templateBookmark, $managedRange) } catch {} }
  $after = $document.Range($contentEnd, $contentEnd); $after.InsertParagraphAfter()
  $operationData.wordReplacement = [pscustomobject]@{ linkId = $linkId; replaced = [bool]$destination.replace; start = $start; replaceEnd = $(if ($null -ne $replaceEnd) { $replaceEnd } else { $start }); inlineShapesBefore = $inlineShapesBefore; deletedInlineShapes = $deletedInlineShapes; inlineShapesAfter = [int]$document.InlineShapes.Count }
  $script:linkedObjects += [pscustomobject]@{ version = 1; target = 'word'; linkId = $linkId; locator = 'bookmark:' + [string]$destination.bookmark; bookmark = [string]$destination.bookmark; source = $(if ($linkedItem) { try { [string]$linkedItem.LinkFormat.SourceFullName } catch { $_filePath } } else { $_filePath }); sourceType = $sourceType; sourceName = $sourceName; range = $rangeAddress; linked = $linked; managed = $true }
}
function Find-PresentationSlide($targetPresentation, [string]$linkId, [int]$requestedSlideId, [int]$requestedSlideIndex) {
  foreach ($candidateSlide in $targetPresentation.Slides) { try { if ([string]$candidateSlide.Tags.Item('WENGGE_LINK_ID') -eq $linkId) { return [pscustomobject]@{ slide = $candidateSlide; managed = $true; created = $false } } } catch {} }
  if ($requestedSlideId -gt 0) { foreach ($candidateSlide in $targetPresentation.Slides) { if ([int]$candidateSlide.SlideID -eq $requestedSlideId) { return [pscustomobject]@{ slide = $candidateSlide; managed = $false; created = $false } } } }
  if ($requestedSlideIndex -gt 0 -and $requestedSlideIndex -le $targetPresentation.Slides.Count) { return [pscustomobject]@{ slide = $targetPresentation.Slides.Item($requestedSlideIndex); managed = $false; created = $false } }
  return [pscustomobject]@{ slide = $targetPresentation.Slides.Add($targetPresentation.Slides.Count + 1, 12); managed = $true; created = $true }
}
function Add-ExcelContentToPresentation($targetPresentation, $sourceSheet, [string]$rangeAddress, [string]$sourceType, [string]$sourceName, [string]$title, [bool]$linked, [string]$linkId, [int]$requestedSlideId = 0, [int]$requestedSlideIndex = 0, [string]$targetShapeName = '') {
  $slideInfo = Find-PresentationSlide $targetPresentation $linkId $requestedSlideId $requestedSlideIndex
  $slide = $slideInfo.slide
  $existingShape = $null
  foreach ($candidateShape in $slide.Shapes) { try { if ([string]$candidateShape.Tags.Item('WENGGE_LINK_ID') -eq $linkId) { $existingShape = $candidateShape; break } } catch {} }
  if ($null -eq $existingShape -and $targetShapeName) { try { $existingShape = $slide.Shapes.Item($targetShapeName) } catch {} }
  $geometry = $null
  if ($null -ne $existingShape) {
    $geometry = [pscustomobject]@{
      left = [double]$existingShape.Left
      top = [double]$existingShape.Top
      width = [double]$existingShape.Width
      height = [double]$existingShape.Height
      rotation = $(try { [double]$existingShape.Rotation } catch { 0 })
      lockAspectRatio = $(try { [int]$existingShape.LockAspectRatio } catch { 0 })
      zOrder = $(try { [int]$existingShape.ZOrderPosition } catch { 0 })
    }
    $existingShape.Delete()
  }
  if ($title) {
    $titleShape = $null
    foreach ($candidateShape in $slide.Shapes) { try { if ([string]$candidateShape.Tags.Item('WENGGE_MANAGED_TITLE') -eq $linkId) { $titleShape = $candidateShape; break } } catch {} }
    if ($null -eq $titleShape) { $titleShape = $slide.Shapes.AddTextbox(1, 30, 20, $targetPresentation.PageSetup.SlideWidth - 60, 40); Set-ObjectTag $titleShape 'WENGGE_MANAGED_TITLE' $linkId }
    $titleShape.TextFrame.TextRange.Text = $title
    $titleShape.TextFrame.TextRange.Font.Size = 24
    $titleShape.TextFrame.TextRange.Font.Bold = $true
  }
  $beforeShapeCount = [int]$slide.Shapes.Count
  $beforeShapeIds = @()
  foreach ($candidateShape in $slide.Shapes) { try { $beforeShapeIds += [int]$candidateShape.Id } catch {} }
  $pastedRange = if ($linked) {
    Invoke-OfficeClipboardPaste { Copy-ExcelContent $sourceSheet $rangeAddress $sourceType $sourceName $false } { $slide.Shapes.PasteSpecial(10, 0, '', 0, '', -1) }
  } else {
    Invoke-OfficeClipboardPaste { Copy-ExcelContent $sourceSheet $rangeAddress $sourceType $sourceName $true } { $slide.Shapes.PasteSpecial(2) }
  }
  if ($slide.Shapes.Count -le $beforeShapeCount) { throw 'PowerPoint 未创建 Excel 内容对象' }
  $shapeHolder = [pscustomobject]@{ value = $null }
  $createdShapeCount = 0
  foreach ($candidateShape in $slide.Shapes) {
    try {
      if ([int]$candidateShape.Id -notin $beforeShapeIds) {
        $createdShapeCount++
        $shapeHolder.value = $candidateShape
      }
    } catch {}
  }
  $shape = if ($createdShapeCount -eq 1) { $shapeHolder.value } else { $null }
  if ($null -eq $shape) {
    $afterShapeDetails = @()
    foreach ($candidateShape in $slide.Shapes) {
      try { $afterShapeDetails += [pscustomobject]@{ id = [int]$candidateShape.Id; name = [string]$candidateShape.Name; type = [int]$candidateShape.Type } } catch {}
    }
    $pasteTypes = @()
    foreach ($candidatePaste in @($pastedRange)) { try { $pasteTypes += [string]$candidatePaste.GetType().FullName } catch {} }
    $diagnostic = [pscustomobject]@{ beforeCount = $beforeShapeCount; afterCount = [int]$slide.Shapes.Count; beforeIds = $beforeShapeIds; after = $afterShapeDetails; createdShapeCount = $createdShapeCount; pasteTypes = $pasteTypes } | ConvertTo-Json -Depth 5 -Compress
    throw "无法定位刚插入的 PowerPoint 内容对象: $diagnostic"
  }
  $shape.Name = 'Wengge Link ' + (Get-LinkKey $linkId)
  Set-ObjectTag $shape 'WENGGE_LINK_ID' $linkId
  Set-ObjectTag $shape 'WENGGE_SOURCE_PATH' $_filePath
  Set-ObjectTag $shape 'WENGGE_SOURCE_TYPE' $sourceType
  Set-ObjectTag $shape 'WENGGE_SOURCE_NAME' $sourceName
  Set-ObjectTag $shape 'WENGGE_SOURCE_RANGE' $rangeAddress
  Set-ObjectTag $shape 'WENGGE_LINKED' $(if ($linked) { 'true' } else { 'false' })
  Set-ObjectTag $shape 'WENGGE_MANAGED' 'true'
  if ([bool]$slideInfo.managed) { Set-ObjectTag $slide 'WENGGE_MANAGED_SLIDE' 'true'; Set-ObjectTag $slide 'WENGGE_LINK_ID' $linkId }
  $maxWidth = $targetPresentation.PageSetup.SlideWidth - 60
  $maxHeight = $targetPresentation.PageSetup.SlideHeight - 100
  if ($shape.Width -gt $maxWidth) { $shape.Width = $maxWidth }
  if ($shape.Height -gt $maxHeight) { $shape.Height = $maxHeight }
  $shape.Left = ($targetPresentation.PageSetup.SlideWidth - $shape.Width) / 2
  $shape.Top = 80 + (($maxHeight - $shape.Height) / 2)
  if ($null -ne $geometry) {
    try { $shape.LockAspectRatio = 0 } catch {}
    $shape.Left = $geometry.left; $shape.Top = $geometry.top; $shape.Width = $geometry.width; $shape.Height = $geometry.height
    try { $shape.Rotation = $geometry.rotation } catch {}
    try { $shape.LockAspectRatio = $geometry.lockAspectRatio } catch {}
    try { while ($geometry.zOrder -gt 0 -and [int]$shape.ZOrderPosition -gt [int]$geometry.zOrder) { $shape.ZOrder(3) } } catch {}
  }
  $script:linkedObjects += [pscustomobject]@{ version = 1; target = 'presentation'; linkId = $linkId; locator = 'shape:' + [int]$slide.SlideID + '/' + [Uri]::EscapeDataString([string]$shape.Name); slideId = [int]$slide.SlideID; slideIndex = [int]$slide.SlideIndex; shapeName = [string]$shape.Name; source = $(if ($linked) { try { [string]$shape.LinkFormat.SourceFullName } catch { $_filePath } } else { $_filePath }); sourceType = $sourceType; sourceName = $sourceName; range = $rangeAddress; linked = $linked; managed = $true; managedSlide = [bool]$slideInfo.managed }
}
try {
  $excelProcessIdsBefore = @(Get-Process -Name 'EXCEL', 'wps' -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
  $wordProcessIdsBefore = @(Get-Process -Name 'WINWORD', 'wps' -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
  $presentationProcessIdsBefore = @(Get-Process -Name 'POWERPNT', 'wpp', 'wps' -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
  $sourceInstanceId = if ($actionParams.sourceInstanceId) { [string]$actionParams.sourceInstanceId } elseif ($actionParams.instanceId) { [string]$actionParams.instanceId } else { '' }
  $sourceHandle = if ($sourceInstanceId) { Find-OfficeDocumentHandle 'excel' $_filePath $sourceInstanceId } else { $null }
  if ($null -ne $sourceHandle) { $excel = $sourceHandle.application; $workbook = $sourceHandle.document }
  else { $excel = New-OfficeComObject ${progIdsLiteral(excelProgIds)}; $excelCreatedApp = $true }
  $excelOwnedProcessId = Get-OwnedOfficeProcessId $excel @('EXCEL', 'wps') $excelProcessIdsBefore
  Register-OwnedOfficeProcess $excelOwnedProcessId
  $excel.Visible = $true
  $wantedWorkbookPath = [System.IO.Path]::GetFullPath($_filePath)
  foreach ($candidateWorkbook in $excel.Workbooks) {
    try {
      if ([System.IO.Path]::GetFullPath([string]$candidateWorkbook.FullName) -ieq $wantedWorkbookPath) {
        $workbook = $candidateWorkbook
        break
      }
    } catch {}
  }
  if ($null -eq $workbook) {
    $workbook = $excel.Workbooks.Open($_filePath)
    $openedWorkbook = $true
  }
  $sheet = $workbook.Worksheets.Item($_sheetName)
  $sourceRange = $sheet.Range($_rangeAddress)
  $overwrite = $actionParams.overwrite -eq $true
  $null = . {
${crossOperationScript(input.operation, { word: wordProgIds, presentation: presentationProgIds })}
  }
  foreach ($replaced in $replacedOutputs) {
    if (Test-Path -LiteralPath $replaced.backup) { Remove-Item -LiteralPath $replaced.backup -Force }
  }
  $replacedOutputs = @()
  $operationData.linkedObjects = $linkedObjects
  [pscustomobject]@{ outputPath = $_outputPath; changes = $changes; data = $operationData } | ConvertTo-Json -Depth 8 -Compress
} catch {
  foreach ($published in $publishedOutputs) {
    if ($published -and (Test-Path -LiteralPath $published)) { Remove-Item -LiteralPath $published -Force -ErrorAction SilentlyContinue }
  }
  foreach ($replaced in $replacedOutputs) {
    if (Test-Path -LiteralPath $replaced.backup) { Move-Item -LiteralPath $replaced.backup -Destination $replaced.destination -Force -ErrorAction SilentlyContinue }
  }
  foreach ($temporary in $tempOutputs) { if ($temporary -and (Test-Path -LiteralPath $temporary)) { Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue } }
  throw
} finally {
  if ($null -ne $wordDoc) { if ($openedWordDocument -or $createdWordDocument) { try { $wordDoc.Close($false) } catch {} }; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wordDoc) } catch {} }
  if ($null -ne $presentation) { if ($openedPresentationDocument -or $createdPresentationDocument) { try { $presentation.Close() } catch {} }; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) } catch {} }
  if ($null -ne $workbook) { if ($openedWorkbook) { try { $workbook.Close($false) } catch {} }; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) } catch {} }
  if ($null -ne $word) { if ($wordCreatedApp) { try { $word.Quit() } catch {} }; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch {} }
  if ($null -ne $powerPoint) { if ($presentationCreatedApp) { try { $powerPoint.Quit() } catch {} }; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint) } catch {} }
  if ($null -ne $excel) { if ($excelCreatedApp) { try { $excel.Quit() } catch {} }; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {} }
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
  Stop-OwnedOfficeProcess $wordOwnedProcessId
  Stop-OwnedOfficeProcess $presentationOwnedProcessId
  Stop-OwnedOfficeProcess $excelOwnedProcessId
  if ($env:WENGGE_MANAGED_PROCESS_ID_FILE) { try { Remove-Item -LiteralPath $env:WENGGE_MANAGED_PROCESS_ID_FILE -Force -ErrorAction SilentlyContinue } catch {} }
}
`;
}

function crossOperationScript(
  operation: string,
  progIds: { word: string[]; presentation: string[] },
): string {
  switch (operation) {
    case "exportRangeToWord":
      return `
  $updateExisting = $actionParams.updateExisting -eq $true
  $wordInstanceId = if ($actionParams.wordInstanceId) { [string]$actionParams.wordInstanceId } else { '' }
  $wordHandle = if ($updateExisting -and $wordInstanceId) { Find-OfficeDocumentHandle 'word' $_outputPath $wordInstanceId } else { $null }
  if ($null -ne $wordHandle) { $word = $wordHandle.application; $wordDoc = $wordHandle.document }
  else { $word = New-OfficeComObject ${progIdsLiteral(progIds.word)}; $wordCreatedApp = $true }
  $wordOwnedProcessId = Get-OwnedOfficeProcessId $word @('WINWORD', 'wps') $wordProcessIdsBefore
  Register-OwnedOfficeProcess $wordOwnedProcessId
  $word.Visible = $true
  if ($null -eq $wordDoc -and $updateExisting -and [IO.File]::Exists($_outputPath)) {
    $wantedWordPath = [IO.Path]::GetFullPath($_outputPath)
    foreach ($candidate in $word.Documents) { try { if ([IO.Path]::GetFullPath([string]$candidate.FullName) -ieq $wantedWordPath) { $wordDoc = $candidate; break } } catch {} }
    if ($null -eq $wordDoc) { $wordDoc = $word.Documents.Open($_outputPath); $openedWordDocument = $true }
  } elseif ($null -eq $wordDoc) { $wordDoc = $word.Documents.Add(); $createdWordDocument = $true }
  $sourceType = if ($actionParams.sourceType) { [string]$actionParams.sourceType } else { 'range' }
  $sourceName = if ($actionParams.chartName) { [string]$actionParams.chartName } elseif ($actionParams.sourceName) { [string]$actionParams.sourceName } else { '' }
  $linkId = if ($actionParams.linkId) { [string]$actionParams.linkId } else { [Guid]::NewGuid().ToString('N') }
  Add-ExcelContentToWord $wordDoc $sheet $_rangeAddress $sourceType $sourceName ([string]$actionParams.title) ($actionParams.linked -eq $true) $linkId ([string]$actionParams.bookmark) ([string]$actionParams.contentControlTag)
  $existingWordIds = @(Get-WordManagedIds $wordDoc)
  Set-WordManagedIds $wordDoc @($existingWordIds + @($linkId))
  if ($updateExisting -and [IO.File]::Exists($_outputPath) -and -not $createdWordDocument) {
    $wordDoc.Save()
  } else {
    $temporary = New-AtomicOutputPath $_outputPath
    $tempOutputs += $temporary
    $wordDoc.SaveAs2($temporary)
    $wordDoc.Close($false); $wordDoc = $null; $createdWordDocument = $false
    Publish-AtomicOutput $temporary $_outputPath $overwrite
    $tempOutputs = @()
  }
  $changes += [pscustomobject]@{ kind = 'cross-office-export'; target = $_outputPath; detail = $(if ($updateExisting) { '已按 linkId 增量更新 Word 受管区域' } elseif ($actionParams.linked -eq $true) { '已将 Excel 内容链接到 Word，可原位刷新' } else { '已将 Excel 内容导出到 Word' }) }
`;
    case "exportRangeToPresentation":
      return `
  $updateExisting = $actionParams.updateExisting -eq $true
  $presentationInstanceId = if ($actionParams.presentationInstanceId) { [string]$actionParams.presentationInstanceId } else { '' }
  $presentationHandle = if ($updateExisting -and $presentationInstanceId) { Find-OfficeDocumentHandle 'presentation' $_outputPath $presentationInstanceId } else { $null }
  if ($null -ne $presentationHandle) { $powerPoint = $presentationHandle.application; $presentation = $presentationHandle.document }
  else { $powerPoint = New-OfficeComObject ${progIdsLiteral(progIds.presentation)}; $presentationCreatedApp = $true }
  $presentationOwnedProcessId = Get-OwnedOfficeProcessId $powerPoint @('POWERPNT', 'wpp', 'wps') $presentationProcessIdsBefore
  Register-OwnedOfficeProcess $presentationOwnedProcessId
  $powerPoint.Visible = -1
  if ($null -eq $presentation -and $updateExisting -and [IO.File]::Exists($_outputPath)) {
    $wantedPresentationPath = [IO.Path]::GetFullPath($_outputPath)
    foreach ($candidate in $powerPoint.Presentations) { try { if ([IO.Path]::GetFullPath([string]$candidate.FullName) -ieq $wantedPresentationPath) { $presentation = $candidate; break } } catch {} }
    if ($null -eq $presentation) { $presentation = $powerPoint.Presentations.Open($_outputPath); $openedPresentationDocument = $true }
  } elseif ($null -eq $presentation) { $presentation = $powerPoint.Presentations.Add(); $createdPresentationDocument = $true }
  $sourceType = if ($actionParams.sourceType) { [string]$actionParams.sourceType } else { 'range' }
  $sourceName = if ($actionParams.chartName) { [string]$actionParams.chartName } elseif ($actionParams.sourceName) { [string]$actionParams.sourceName } else { '' }
  $linkId = if ($actionParams.linkId) { [string]$actionParams.linkId } else { [Guid]::NewGuid().ToString('N') }
  Add-ExcelContentToPresentation $presentation $sheet $_rangeAddress $sourceType $sourceName ([string]$actionParams.title) ($actionParams.linked -eq $true) $linkId ([int]$actionParams.slideId) ([int]$actionParams.slideIndex) ([string]$actionParams.shapeName)
  $existingPresentationIds = @(Get-PresentationManagedIds $presentation)
  Set-PresentationManagedIds $presentation @($existingPresentationIds + @($linkId))
  if ($updateExisting -and [IO.File]::Exists($_outputPath) -and -not $createdPresentationDocument) {
    $presentation.Save()
  } else {
    $temporary = New-AtomicOutputPath $_outputPath
    $tempOutputs += $temporary
    $presentation.SaveAs($temporary)
    $presentation.Close(); $presentation = $null; $createdPresentationDocument = $false
    Publish-AtomicOutput $temporary $_outputPath $overwrite
    $tempOutputs = @()
  }
  $changes += [pscustomobject]@{ kind = 'cross-office-export'; target = $_outputPath; detail = $(if ($updateExisting) { '已按 linkId 增量更新 PowerPoint 受管对象' } elseif ($actionParams.linked -eq $true) { '已将 Excel 内容链接到 PowerPoint，可原位刷新' } else { '已将 Excel 内容导出到 PowerPoint' }) }
`;
    case "buildReportPackage":
      return `
  $outputDirectory = if ($actionParams.outputDirectory) { [string]$actionParams.outputDirectory } else { [System.IO.Path]::GetDirectoryName($_outputPath) }
  $baseName = if ($actionParams.baseName) { [string]$actionParams.baseName } else { [System.IO.Path]::GetFileNameWithoutExtension($_filePath) + '-报告' }
  $wordOutput = if ($actionParams.wordOutputPath) { [string]$actionParams.wordOutputPath } else { [System.IO.Path]::Combine($outputDirectory, $baseName + '.docx') }
  $presentationOutput = if ($actionParams.presentationOutputPath) { [string]$actionParams.presentationOutputPath } else { [System.IO.Path]::Combine($outputDirectory, $baseName + '.pptx') }
  $updateExisting = $actionParams.updateExisting -eq $true
  if ((Test-Path -LiteralPath $wordOutput) -and -not $overwrite -and -not $updateExisting) { throw "输出文件已存在: $wordOutput" }
  if ((Test-Path -LiteralPath $presentationOutput) -and -not $overwrite -and -not $updateExisting) { throw "输出文件已存在: $presentationOutput" }
  $wordInstanceId = if ($actionParams.wordInstanceId) { [string]$actionParams.wordInstanceId } else { '' }
  $wordHandle = if ($updateExisting -and $wordInstanceId) { Find-OfficeDocumentHandle 'word' $wordOutput $wordInstanceId } else { $null }
  if ($null -ne $wordHandle) { $word = $wordHandle.application; $wordDoc = $wordHandle.document }
  else { $word = New-OfficeComObject ${progIdsLiteral(progIds.word)}; $wordCreatedApp = $true }
  $wordOwnedProcessId = Get-OwnedOfficeProcessId $word @('WINWORD', 'wps') $wordProcessIdsBefore
  Register-OwnedOfficeProcess $wordOwnedProcessId
  $word.Visible = $true
  if ($null -eq $wordDoc -and $updateExisting -and [IO.File]::Exists($wordOutput)) {
    $wantedWordPath = [IO.Path]::GetFullPath($wordOutput)
    foreach ($candidate in $word.Documents) { try { if ([IO.Path]::GetFullPath([string]$candidate.FullName) -ieq $wantedWordPath) { $wordDoc = $candidate; break } } catch {} }
    if ($null -eq $wordDoc) { $wordDoc = $word.Documents.Open($wordOutput); $openedWordDocument = $true }
  } elseif ($null -eq $wordDoc) { $wordDoc = $word.Documents.Add(); $createdWordDocument = $true }
  $presentationInstanceId = if ($actionParams.presentationInstanceId) { [string]$actionParams.presentationInstanceId } else { '' }
  $presentationHandle = if ($updateExisting -and $presentationInstanceId) { Find-OfficeDocumentHandle 'presentation' $presentationOutput $presentationInstanceId } else { $null }
  if ($null -ne $presentationHandle) { $powerPoint = $presentationHandle.application; $presentation = $presentationHandle.document }
  else { $powerPoint = New-OfficeComObject ${progIdsLiteral(progIds.presentation)}; $presentationCreatedApp = $true }
  $presentationOwnedProcessId = Get-OwnedOfficeProcessId $powerPoint @('POWERPNT', 'wpp', 'wps') $presentationProcessIdsBefore
  Register-OwnedOfficeProcess $presentationOwnedProcessId
  $powerPoint.Visible = -1
  if ($null -eq $presentation -and $updateExisting -and [IO.File]::Exists($presentationOutput)) {
    $wantedPresentationPath = [IO.Path]::GetFullPath($presentationOutput)
    foreach ($candidate in $powerPoint.Presentations) { try { if ([IO.Path]::GetFullPath([string]$candidate.FullName) -ieq $wantedPresentationPath) { $presentation = $candidate; break } } catch {} }
    if ($null -eq $presentation) { $presentation = $powerPoint.Presentations.Open($presentationOutput); $openedPresentationDocument = $true }
  } elseif ($null -eq $presentation) { $presentation = $powerPoint.Presentations.Add(); $createdPresentationDocument = $true }
  $sections = @($actionParams.sections)
  if ($sections.Count -eq 0) { $sections = @([pscustomobject]@{ linkId = [string]$actionParams.linkId; sheetName = $_sheetName; range = $_rangeAddress; title = [string]$actionParams.title }) }
  $requestedWordIds = @(); $requestedPresentationIds = @()
  foreach ($section in $sections) {
    $sectionSheet = if ($section.sheetName) { $workbook.Worksheets.Item([string]$section.sheetName) } else { $sheet }
    $sectionRangeAddress = if ($section.range) { [string]$section.range } else { $_rangeAddress }
    $sectionSourceType = if ($section.sourceType) { [string]$section.sourceType } else { 'range' }
    $sectionSourceName = if ($section.chartName) { [string]$section.chartName } elseif ($section.sourceName) { [string]$section.sourceName } else { '' }
    $sectionLinked = if ($null -ne $section.linked) { $section.linked -eq $true } else { $actionParams.linked -eq $true }
    $sectionLinkId = if ($section.linkId) { [string]$section.linkId } else { [Guid]::NewGuid().ToString('N') }
    $wordLinkId = $sectionLinkId + '_word'; $presentationLinkId = $sectionLinkId + '_ppt'
    $requestedWordIds += $wordLinkId; $requestedPresentationIds += $presentationLinkId
    Add-ExcelContentToWord $wordDoc $sectionSheet $sectionRangeAddress $sectionSourceType $sectionSourceName ([string]$section.title) $sectionLinked $wordLinkId ([string]$section.bookmark) ([string]$section.contentControlTag)
    Add-ExcelContentToPresentation $presentation $sectionSheet $sectionRangeAddress $sectionSourceType $sectionSourceName ([string]$section.title) $sectionLinked $presentationLinkId ([int]$section.slideId) ([int]$section.slideIndex) ([string]$section.shapeName)
  }
  if ($updateExisting) {
    foreach ($oldId in @(Get-WordManagedIds $wordDoc)) {
      if ($oldId -notin $requestedWordIds) { $bookmarkName = 'WgLink_' + (Get-LinkKey ([string]$oldId)); if ($wordDoc.Bookmarks.Exists($bookmarkName)) { $range = $wordDoc.Bookmarks.Item($bookmarkName).Range; $wordDoc.Bookmarks.Item($bookmarkName).Delete(); $range.Delete() }; Remove-WordLinkMetadata $wordDoc $bookmarkName }
    }
    for ($slideIndex = $presentation.Slides.Count; $slideIndex -ge 1; $slideIndex--) {
      $candidateSlide = $presentation.Slides.Item($slideIndex); $managed = try { [string]$candidateSlide.Tags.Item('WENGGE_MANAGED_SLIDE') } catch { '' }; $oldId = try { [string]$candidateSlide.Tags.Item('WENGGE_LINK_ID') } catch { '' }
      if ($managed -eq 'true' -and $oldId -and $oldId -notin $requestedPresentationIds) { $candidateSlide.Delete() }
    }
  }
  Set-WordManagedIds $wordDoc $requestedWordIds
  Set-PresentationManagedIds $presentation $requestedPresentationIds
  if ($createdWordDocument) { $wordTemporary = New-AtomicOutputPath $wordOutput; $tempOutputs += $wordTemporary; $wordDoc.SaveAs2($wordTemporary); $wordDoc.Close($false); $wordDoc = $null; $createdWordDocument = $false; Publish-AtomicOutput $wordTemporary $wordOutput $overwrite }
  else { $wordDoc.Save() }
  if ($createdPresentationDocument) { $presentationTemporary = New-AtomicOutputPath $presentationOutput; $tempOutputs += $presentationTemporary; $presentation.SaveAs($presentationTemporary); $presentation.Close(); $presentation = $null; $createdPresentationDocument = $false; Publish-AtomicOutput $presentationTemporary $presentationOutput $overwrite }
  else { $presentation.Save() }
  $tempOutputs = @()
  $_outputPath = $outputDirectory
  $operationData.manifest = [pscustomobject]@{ version = 1; word = $requestedWordIds; presentation = $requestedPresentationIds; links = $linkedObjects }
  $changes += [pscustomobject]@{ kind = 'report-package'; target = $wordOutput; detail = $(if ($updateExisting) { '已增量同步 Word 受管章节并保留非受管内容' } else { '已生成 Word 报告' }) }
  $changes += [pscustomobject]@{ kind = 'report-package'; target = $presentationOutput; detail = $(if ($updateExisting) { '已增量同步 PowerPoint 受管幻灯片并保留非受管内容' } else { '已生成 PowerPoint 汇报' }) }
`;
    default:
      return "  throw \"不支持的跨 Office 操作: $_operation\"";
  }
}

function validateIncrementalLinkIds(input: OfficeActionInput): void {
  if (input.params?.updateExisting !== true) return;
  if (input.operation !== "buildReportPackage") {
    if (!hasLinkId(input.params?.linkId)) throw new Error(`${input.operation} 增量更新需要 params.linkId`);
    return;
  }
  const sections = input.params?.sections;
  if (!Array.isArray(sections) || sections.length === 0) {
    if (!hasLinkId(input.params?.linkId)) throw new Error("buildReportPackage 增量更新需要 params.linkId 或每个 section.linkId");
    return;
  }
  const missing = sections.findIndex((section) => !section || typeof section !== "object" || !hasLinkId((section as Record<string, unknown>).linkId));
  if (missing >= 0) throw new Error(`buildReportPackage 增量更新的第 ${missing + 1} 个 section 缺少稳定 linkId`);
}

function hasLinkId(value: unknown): value is string {
  return typeof value === "string" && value.trim().length > 0;
}
