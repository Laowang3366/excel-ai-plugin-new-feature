const LINKED_CONTENT_OPERATIONS = new Set([
  "inspectLinkedOfficeContent",
  "refreshLinkedOfficeContent",
  "relinkLinkedOfficeContent",
]);

export function buildWordLinkedContentOperationScript(operation: string): string | undefined {
  if (!LINKED_CONTENT_OPERATIONS.has(operation)) return undefined;
  return String.raw`
  $links = @(); $updated = 0; $relinked = 0; $failures = @(); $matched = 0
  $linkIdFilter = if ($actionParams.linkId) { [string]$actionParams.linkId } else { '' }
  $newSourcePath = if ($actionParams.sourcePath) { [IO.Path]::GetFullPath([string]$actionParams.sourcePath) } else { '' }
  if ($_operation -eq 'relinkLinkedOfficeContent' -and (-not $linkIdFilter -or -not $newSourcePath)) { throw '重绑链接需要 params.linkId 和 params.sourcePath' }
  function Get-RelinkSource([string]$oldSource, [string]$newPath) { $suffix = ''; $separator = $oldSource.IndexOf('!'); if ($separator -ge 0) { $oldPath = $oldSource.Substring(0, $separator); $suffix = $oldSource.Substring($separator); $oldName = [IO.Path]::GetFileName($oldPath); $newName = [IO.Path]::GetFileName($newPath); if ($oldName -and $newName) { $suffix = $suffix.Replace('[' + $oldName + ']', '[' + $newName + ']') } }; return $newPath + $suffix }
  function Set-LinkedContentTag($item, [string]$name, [string]$value) { try { $item.Tags.Delete($name) } catch {}; try { $item.Tags.Add($name, $value) } catch {} }
  function Find-ContainingWordBookmark($document, $range) {
    foreach ($candidate in $document.Bookmarks) {
      try { if ([int]$candidate.Range.Start -le [int]$range.Start -and [int]$candidate.Range.End -ge [int]$range.End) { return [string]$candidate.Name } } catch {}
    }
    return ''
  }
  function Get-WordLinkKey([string]$linkId) {
    $safe = $linkId -replace '[^A-Za-z0-9_]', '_'; if ($safe.Length -gt 16) { $safe = $safe.Substring(0, 16) }
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $hash = [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($linkId))).Replace('-', '').Substring(0, 10) } finally { $sha.Dispose() }
    return $safe + '_' + $hash
  }
  $managedIds = @()
  try { $managedIds = @(([string]$doc.Variables.Item('WENGGE_MANAGED_LINK_IDS').Value | ConvertFrom-Json)) } catch { $managedIds = @() }
  $canMapManagedIdsByOrder = $managedIds.Count -gt 0 -and [int]$doc.InlineShapes.Count -eq $managedIds.Count
  $managedLinkOrdinal = 0
  $linkCandidates = @()
${linkedExcelSetup(operation)}
  try {
  for ($index = 1; $index -le $doc.InlineShapes.Count; $index++) {
    $item = $doc.InlineShapes.Item($index)
    $source = try { [string]$item.LinkFormat.SourceFullName } catch { '' }
    if (-not $source) { continue }
    $managedLinkOrdinal++
    $bookmark = Find-ContainingWordBookmark $doc $item.Range
    $manifest = $null; $linkId = ''
    try { $alternativeText = [string]$item.AlternativeText; if ($alternativeText.StartsWith('WENGGE_MANIFEST:')) { $manifest = $alternativeText.Substring(16) | ConvertFrom-Json; $linkId = [string]$manifest.linkId } } catch {}
    if (-not $linkId -and $bookmark) { try { $manifest = ([string]$doc.Variables.Item('WENGGE_META_' + $bookmark).Value | ConvertFrom-Json); $linkId = [string]$manifest.linkId } catch {} }
    if (-not $linkId -and $bookmark) { foreach ($managedId in $managedIds) { if ($bookmark -eq ('WgLink_' + (Get-WordLinkKey ([string]$managedId)))) { $linkId = [string]$managedId; break } } }
    if (-not $linkId -and $canMapManagedIdsByOrder) { $linkId = [string]$managedIds[$managedLinkOrdinal - 1]; $manifest = [pscustomobject]@{ version = 1; linkId = $linkId; managed = $true } }
    $linkCandidates += [pscustomobject]@{ index = $index; bookmark = $bookmark; linkId = $linkId; source = $source }
    if ($linkIdFilter -and $linkId -ne $linkIdFilter) { continue }
    $matched++
    if ($_operation -eq 'relinkLinkedOfficeContent') { try { Open-LinkedExcelSource $newSourcePath; $item.LinkFormat.SourceFullName = Get-RelinkSource $source $newSourcePath; $source = [string]$item.LinkFormat.SourceFullName; if ($null -ne $manifest -and $bookmark) { $manifest.source = $newSourcePath; $manifestJson = $manifest | ConvertTo-Json -Compress; try { $doc.Variables.Item('WENGGE_META_' + $bookmark).Value = $manifestJson } catch {}; try { $item.AlternativeText = 'WENGGE_MANIFEST:' + $manifestJson } catch {} }; $relinked++ } catch { $failures += [pscustomobject]@{ kind = 'inlineShape'; index = $index; linkId = $linkId; error = $_.Exception.Message } } }
    if ($_operation -eq 'refreshLinkedOfficeContent') {
      try { Open-LinkedExcelSource $source; $item.LinkFormat.Update(); $updated++ } catch { $failures += [pscustomobject]@{ kind = 'inlineShape'; index = $index; error = $_.Exception.Message } }
    }
    $links += [pscustomobject]@{ version = 1; kind = 'inlineShape'; index = $index; name = $bookmark; linkId = $linkId; source = $source; locator = 'inlineShape:' + $index; managed = $null -ne $manifest; metadata = $manifest }
  }
  for ($index = 1; $index -le $doc.Shapes.Count; $index++) {
    $item = $doc.Shapes.Item($index)
    $source = try { [string]$item.LinkFormat.SourceFullName } catch { '' }
    if (-not $source) { continue }
    $linkId = try { [string]$item.AlternativeText } catch { '' }
    if ($linkIdFilter -and $linkId -ne $linkIdFilter) { continue }
    $matched++
    if ($_operation -eq 'relinkLinkedOfficeContent') { try { $item.LinkFormat.SourceFullName = Get-RelinkSource $source $newSourcePath; $source = [string]$item.LinkFormat.SourceFullName; $relinked++ } catch { $failures += [pscustomobject]@{ kind = 'shape'; index = $index; linkId = $linkId; error = $_.Exception.Message } } }
    if ($_operation -eq 'refreshLinkedOfficeContent') {
      try { Open-LinkedExcelSource $source; $item.LinkFormat.Update(); $updated++ } catch { $failures += [pscustomobject]@{ kind = 'shape'; index = $index; error = $_.Exception.Message } }
    }
    $links += [pscustomobject]@{ version = 1; kind = 'shape'; index = $index; name = [string]$item.Name; linkId = $linkId; source = $source; locator = 'shape:' + [Uri]::EscapeDataString([string]$item.Name) }
  }
  if ($linkIdFilter -and $matched -eq 0) { throw '找不到指定 linkId 的 Word 链接对象: ' + $linkIdFilter + '; candidates=' + ($linkCandidates | ConvertTo-Json -Compress) + '; managedIds=' + ($managedIds | ConvertTo-Json -Compress) }
  $manifestVersion = try { [string]$doc.Variables.Item('WENGGE_MANIFEST_VERSION').Value } catch { '' }
  $operationData.links = $links; $operationData.updated = $updated; $operationData.relinked = $relinked; $operationData.failures = $failures; $operationData.manifest = [pscustomobject]@{ version = $manifestVersion; managedLinkIds = $managedIds }
  if ($_operation -eq 'refreshLinkedOfficeContent') {
    if ($failures.Count -gt 0 -and $updated -eq 0) { throw 'Word 链接内容刷新失败: ' + ($failures | ConvertTo-Json -Compress) }
    $changes += [pscustomobject]@{ kind = 'linked-content-refresh'; target = $_filePath; detail = '已原位刷新 ' + $updated + ' 个 Excel 链接对象' }
  } elseif ($_operation -eq 'relinkLinkedOfficeContent') {
    if ($failures.Count -gt 0 -and $relinked -eq 0) { throw 'Word 链接重绑失败: ' + ($failures | ConvertTo-Json -Compress) }
    $changes += [pscustomobject]@{ kind = 'linked-content-relink'; target = $_filePath; detail = '已重绑 ' + $relinked + ' 个 Word 链接对象' }
  } else { $changes = @() }
  } finally {
${linkedExcelCleanup(operation)}
  }
`;
}

export function buildPresentationLinkedContentOperationScript(operation: string): string | undefined {
  if (!LINKED_CONTENT_OPERATIONS.has(operation)) return undefined;
  return String.raw`
  $links = @(); $updated = 0; $relinked = 0; $failures = @(); $matched = 0
  $linkIdFilter = if ($actionParams.linkId) { [string]$actionParams.linkId } else { '' }
  $newSourcePath = if ($actionParams.sourcePath) { [IO.Path]::GetFullPath([string]$actionParams.sourcePath) } else { '' }
  if ($_operation -eq 'relinkLinkedOfficeContent' -and (-not $linkIdFilter -or -not $newSourcePath)) { throw '重绑链接需要 params.linkId 和 params.sourcePath' }
  function Get-RelinkSource([string]$oldSource, [string]$newPath) { $suffix = ''; $separator = $oldSource.IndexOf('!'); if ($separator -ge 0) { $oldPath = $oldSource.Substring(0, $separator); $suffix = $oldSource.Substring($separator); $oldName = [IO.Path]::GetFileName($oldPath); $newName = [IO.Path]::GetFileName($newPath); if ($oldName -and $newName) { $suffix = $suffix.Replace('[' + $oldName + ']', '[' + $newName + ']') } }; return $newPath + $suffix }
  function Set-LinkedContentTag($item, [string]$name, [string]$value) { try { $item.Tags.Delete($name) } catch {}; try { $item.Tags.Add($name, $value) } catch {} }
${linkedExcelSetup(operation)}
  try {
  foreach ($targetSlide in $pres.Slides) {
    foreach ($item in $targetSlide.Shapes) {
      $source = try { [string]$item.LinkFormat.SourceFullName } catch { '' }
      if (-not $source) { continue }
      $linkId = try { [string]$item.Tags.Item('WENGGE_LINK_ID') } catch { '' }
      if ($linkIdFilter -and $linkId -ne $linkIdFilter) { continue }
      $matched++
      if ($_operation -eq 'relinkLinkedOfficeContent') { try { Open-LinkedExcelSource $newSourcePath; $item.LinkFormat.SourceFullName = Get-RelinkSource $source $newSourcePath; Set-LinkedContentTag $item 'WENGGE_SOURCE_PATH' $newSourcePath; $source = [string]$item.LinkFormat.SourceFullName; $relinked++ } catch { $failures += [pscustomobject]@{ slideIndex = [int]$targetSlide.SlideIndex; name = [string]$item.Name; linkId = $linkId; error = $_.Exception.Message } } }
      if ($_operation -eq 'refreshLinkedOfficeContent') {
        try { Open-LinkedExcelSource $source; $item.LinkFormat.Update(); $updated++ } catch { $failures += [pscustomobject]@{ slideIndex = [int]$targetSlide.SlideIndex; name = [string]$item.Name; error = $_.Exception.Message } }
      }
      $links += [pscustomobject]@{ version = 1; kind = 'shape'; slideId = [int]$targetSlide.SlideID; slideIndex = [int]$targetSlide.SlideIndex; name = [string]$item.Name; linkId = $linkId; source = $source; sourceType = $(try { [string]$item.Tags.Item('WENGGE_SOURCE_TYPE') } catch { '' }); sourceName = $(try { [string]$item.Tags.Item('WENGGE_SOURCE_NAME') } catch { '' }); range = $(try { [string]$item.Tags.Item('WENGGE_SOURCE_RANGE') } catch { '' }); managed = $(try { [string]$item.Tags.Item('WENGGE_MANAGED') -eq 'true' } catch { $false }); locator = 'shape:' + [int]$targetSlide.SlideID + '/' + [Uri]::EscapeDataString([string]$item.Name) }
    }
  }
  if ($linkIdFilter -and $matched -eq 0) { throw '找不到指定 linkId 的演示文稿链接对象: ' + $linkIdFilter }
  $managedIds = @()
  try { $managedIds = @(([string]$pres.Tags.Item('WENGGE_MANAGED_LINK_IDS') | ConvertFrom-Json)) } catch { $managedIds = @() }
  $manifestVersion = try { [string]$pres.Tags.Item('WENGGE_MANIFEST_VERSION') } catch { '' }
  $operationData.links = $links; $operationData.updated = $updated; $operationData.relinked = $relinked; $operationData.failures = $failures; $operationData.manifest = [pscustomobject]@{ version = $manifestVersion; managedLinkIds = $managedIds }
  if ($_operation -eq 'refreshLinkedOfficeContent') {
    if ($failures.Count -gt 0 -and $updated -eq 0) { throw '演示文稿链接内容刷新失败: ' + ($failures | ConvertTo-Json -Compress) }
    $changes += [pscustomobject]@{ kind = 'linked-content-refresh'; target = $_filePath; detail = '已原位刷新 ' + $updated + ' 个 Excel 链接对象' }
  } elseif ($_operation -eq 'relinkLinkedOfficeContent') {
    if ($failures.Count -gt 0 -and $relinked -eq 0) { throw '演示文稿链接重绑失败: ' + ($failures | ConvertTo-Json -Compress) }
    $changes += [pscustomobject]@{ kind = 'linked-content-relink'; target = $_filePath; detail = '已重绑 ' + $relinked + ' 个演示文稿链接对象' }
  } else { $changes = @() }
  } finally {
${linkedExcelCleanup(operation)}
  }
`;
}

function linkedExcelSetup(operation: string): string {
  const processBaseline = String.raw`
  $linkedExcelProcessIdsBefore = @(Get-Process -Name 'EXCEL', 'wps' -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
`;
  if (!['refreshLinkedOfficeContent', 'relinkLinkedOfficeContent'].includes(operation)) return processBaseline;
  return String.raw`${processBaseline}
  $linkedExcel = $null; $linkedExcelCreated = $false; $linkedExcelOwnedProcessId = [uint32]0; $linkedExcelOpenedWorkbooks = @()
  function Resolve-LinkedExcelSource([string]$source) {
    if (Test-Path -LiteralPath $source) { return [IO.Path]::GetFullPath($source) }
    $match = [regex]::Match($source, '(?i)^(.+?\.(xlsx|xlsm|xlsb|xls))(?=!|$)')
    if ($match.Success -and (Test-Path -LiteralPath $match.Groups[1].Value)) { return [IO.Path]::GetFullPath($match.Groups[1].Value) }
    return ''
  }
  function Open-LinkedExcelSource([string]$source) {
    $sourcePath = Resolve-LinkedExcelSource $source
    if (-not $sourcePath) { throw '找不到 Excel 链接源: ' + $source }
    if ($null -eq $script:linkedExcel) {
      foreach ($progId in @('Excel.Application', 'Ket.Application')) {
        foreach ($attempt in 1..3) {
          try { $script:linkedExcel = New-Object -ComObject $progId; $script:linkedExcelCreated = $true; break } catch {}
          Start-Sleep -Milliseconds (200 * $attempt)
        }
        if ($null -ne $script:linkedExcel) { break }
      }
      if ($null -eq $script:linkedExcel) { throw '无法启动 Excel/WPS 表格以刷新链接' }
      $script:linkedExcel.Visible = $false
      $script:linkedExcel.DisplayAlerts = $false
      try {
        if (-not ('WenggeLinkedExcelWindow' -as [type])) { Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class WenggeLinkedExcelWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }' }
        [void][WenggeLinkedExcelWindow]::GetWindowThreadProcessId([IntPtr]$script:linkedExcel.Hwnd, [ref]$script:linkedExcelOwnedProcessId)
        if ([int]$script:linkedExcelOwnedProcessId -in $linkedExcelProcessIdsBefore) { $script:linkedExcelOwnedProcessId = [uint32]0 }
      } catch {}
      if ($script:linkedExcelOwnedProcessId -eq 0) {
        try {
          $createdProcess = Get-Process -Name 'EXCEL', 'wps' -ErrorAction SilentlyContinue | Where-Object { [int]$_.Id -notin $linkedExcelProcessIdsBefore } | Sort-Object StartTime -Descending | Select-Object -First 1
          if ($null -ne $createdProcess) { $script:linkedExcelOwnedProcessId = [uint32]$createdProcess.Id }
        } catch {}
      }
      if ($script:linkedExcelOwnedProcessId -gt 0 -and $env:WENGGE_MANAGED_PROCESS_ID_FILE) {
        try {
          $managedIds = @()
          try { $managedIds = @(([string][IO.File]::ReadAllText($env:WENGGE_MANAGED_PROCESS_ID_FILE) | ConvertFrom-Json)) } catch {}
          $managedIds = @($managedIds + [int]$script:linkedExcelOwnedProcessId | ForEach-Object { [int]$_ } | Where-Object { $_ -gt 0 } | Select-Object -Unique)
          [IO.File]::WriteAllText($env:WENGGE_MANAGED_PROCESS_ID_FILE, ($managedIds | ConvertTo-Json -Compress))
        } catch {}
      }
    }
    foreach ($openWorkbook in $script:linkedExcel.Workbooks) { try { if ([IO.Path]::GetFullPath([string]$openWorkbook.FullName) -ieq $sourcePath) { return } } catch {} }
    $sourceWorkbook = $script:linkedExcel.Workbooks.Open($sourcePath, 0, $true)
    $script:linkedExcelOpenedWorkbooks += $sourceWorkbook
    try { $script:linkedExcel.CalculateFull() } catch {}
  }
`;
}

function linkedExcelCleanup(operation: string): string {
  const cleanupSpawnedExcel = String.raw`
    foreach ($spawnedProcess in @(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue | Where-Object { [int]$_.Id -notin $linkedExcelProcessIdsBefore -and $_.MainWindowHandle -eq 0 })) {
      try { Stop-Process -Id $spawnedProcess.Id -Force -ErrorAction SilentlyContinue } catch {}
    }
`;
  if (!['refreshLinkedOfficeContent', 'relinkLinkedOfficeContent'].includes(operation)) return cleanupSpawnedExcel;
  return String.raw`
    foreach ($sourceWorkbook in @($script:linkedExcelOpenedWorkbooks)) { try { $sourceWorkbook.Close($false) } catch {}; try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($sourceWorkbook) } catch {} }
    if ($null -ne $script:linkedExcel) {
      if ($script:linkedExcelCreated) { try { $script:linkedExcel.Quit() } catch {} }
      try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:linkedExcel) } catch {}
      $script:linkedExcel = $null
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($script:linkedExcelOwnedProcessId -gt 0) {
      try { $process = Get-Process -Id $script:linkedExcelOwnedProcessId -ErrorAction SilentlyContinue; if ($null -ne $process) { [void]$process.WaitForExit(3000) }; $process = Get-Process -Id $script:linkedExcelOwnedProcessId -ErrorAction SilentlyContinue; if ($null -ne $process) { Stop-Process -Id $script:linkedExcelOwnedProcessId -Force -ErrorAction SilentlyContinue } } catch {}
    }
    if ($env:WENGGE_MANAGED_PROCESS_ID_FILE) { try { Remove-Item -LiteralPath $env:WENGGE_MANAGED_PROCESS_ID_FILE -Force -ErrorAction SilentlyContinue } catch {} }
${cleanupSpawnedExcel}
`;
}
