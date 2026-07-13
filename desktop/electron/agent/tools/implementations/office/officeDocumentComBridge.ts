import path from "node:path";

import { safeJsonParse } from "../../../automation/json";
import { executePowerShell, psVar } from "../../../automation/powershell";
import type {
  OfficeDocumentInfo,
  OfficeDocumentManagerBridge,
  OfficeObjectInfo,
} from "../../contracts/office";
import type { OfficeActionApp } from "../../officeCore/types";

export class OfficeDocumentComBridge implements OfficeDocumentManagerBridge {
  async listDocuments(app?: OfficeActionApp): Promise<OfficeDocumentInfo[]> {
    const output = await executePowerShell(`
${officeInstanceDiscoveryScript()}
${psVar("_appFilter", app || "")}
$documents = @(Get-AllOfficeDocumentHandles $_appFilter | ForEach-Object { Get-OfficeDocumentMetadata $_ })
$documents | ConvertTo-Json -Depth 6 -Compress
`);
    const parsed = safeJsonParse<OfficeDocumentInfo | OfficeDocumentInfo[]>(output || "[]", "powershell", "列出 Office 文档");
    return Array.isArray(parsed) ? parsed : parsed ? [parsed] : [];
  }

  async activateDocument(input: {
    app: OfficeActionApp;
    filePath?: string;
    name?: string;
    index?: number;
    instanceId?: string;
  }): Promise<OfficeDocumentInfo> {
    const selector = JSON.stringify({
      filePath: input.filePath,
      name: input.name,
      index: input.index,
      instanceId: input.instanceId,
    });
    const output = await executePowerShell(`
${officeInstanceDiscoveryScript()}
${psVar("_selectorJson", selector)}
$selector = ConvertFrom-Json $_selectorJson
$handle = Resolve-OfficeDocumentHandle '${input.app}' ([string]$selector.filePath) ([string]$selector.instanceId) ([string]$selector.name) ([int]$selector.index)
$handle.application.Visible = ${input.app === "presentation" ? "-1" : "$true"}
$handle.document.Activate()
$metadata = Get-OfficeDocumentMetadata $handle
$metadata.active = $true
$metadata | ConvertTo-Json -Depth 5 -Compress
`);
    return safeJsonParse<OfficeDocumentInfo>(output, "powershell", "激活 Office 文档");
  }

  async listObjects(input: { app: OfficeActionApp; filePath: string; instanceId?: string; kind?: string }): Promise<OfficeObjectInfo[]> {
    const output = await executePowerShell(`
${officeDocumentResolverScript(input.app, input.filePath, input.instanceId)}
${psVar("_kind", input.kind || "")}
$objects = @()
function Add-OfficeObject([string]$kind, [string]$name, [string]$locator, [string]$parent = '', [int]$index = 0, [string]$detail = '', [bool]$selected = $false) {
  if ($_kind -and $_kind -ne $kind) { return }
  $script:objects += [pscustomobject]@{ app = '${input.app}'; documentPath = [string]$target.FullName; instanceId = [string]$handle.instanceId; kind = $kind; name = $name; locator = $locator; parent = $parent; index = $index; detail = $detail; selected = $selected }
}
function Encode-LocatorPart([string]$value) { return [Uri]::EscapeDataString($value) }
${listObjectsScript(input.app)}
$objects | ConvertTo-Json -Depth 6 -Compress
`);
    const parsed = safeJsonParse<OfficeObjectInfo | OfficeObjectInfo[]>(output || "[]", "powershell", "列出 Office 对象");
    return Array.isArray(parsed) ? parsed : parsed ? [parsed] : [];
  }

  async activateObject(input: { app: OfficeActionApp; filePath: string; instanceId?: string; locator: string }): Promise<OfficeObjectInfo> {
    const output = await executePowerShell(`
${officeDocumentResolverScript(input.app, input.filePath, input.instanceId)}
${psVar("_locator", input.locator)}
function Decode-LocatorPart([string]$value) { return [Uri]::UnescapeDataString($value) }
$selectedKind = ''
$selectedName = ''
$selectedParent = ''
$selectedIndex = 0
${activateObjectScript(input.app)}
if (-not $selectedKind) { throw '不支持或找不到 Office 对象 locator: ' + $_locator }
$target.Activate()
[pscustomobject]@{ app = '${input.app}'; documentPath = [string]$target.FullName; instanceId = [string]$handle.instanceId; kind = $selectedKind; name = $selectedName; locator = $_locator; parent = $selectedParent; index = $selectedIndex; selected = $true } | ConvertTo-Json -Depth 5 -Compress
`);
    return safeJsonParse<OfficeObjectInfo>(output, "powershell", "激活 Office 对象");
  }

  async prepareTransaction(filePaths: string[]): Promise<Array<{
    app: OfficeActionApp;
      filePath: string;
      wasDirty: boolean;
      saved: boolean;
      instanceId?: string;
    }>> {
    const normalized = [...new Set(filePaths.map((filePath) => pathKey(filePath)))];
    const output = await executePowerShell(`
${officeInstanceDiscoveryScript()}
${psVar("_pathsJson", JSON.stringify(normalized))}
$wantedPaths = [string[]](ConvertFrom-Json $_pathsJson)
$prepared = @()
foreach ($handle in @(Get-AllOfficeDocumentHandles '')) {
  $fullName = try { [IO.Path]::GetFullPath([string]$handle.document.FullName).ToLowerInvariant() } catch { '' }
  if (-not $fullName -or $fullName -notin $wantedPaths) { continue }
  $wasDirty = try { -not [bool]$handle.document.Saved } catch { $false }
  $saved = $true
  if ($wasDirty) { try { $handle.document.Save() } catch { $saved = $false } }
  $prepared += [pscustomobject]@{ app = [string]$handle.app; filePath = [string]$handle.document.FullName; instanceId = [string]$handle.instanceId; wasDirty = $wasDirty; saved = $saved }
}
$prepared | ConvertTo-Json -Depth 5 -Compress
`);
    const parsed = safeJsonParse<
      { app: OfficeActionApp; filePath: string; instanceId?: string; wasDirty: boolean; saved: boolean }
      | Array<{ app: OfficeActionApp; filePath: string; instanceId?: string; wasDirty: boolean; saved: boolean }>
    >(output || "[]", "powershell", "准备 Office 事务文件");
    const prepared = Array.isArray(parsed) ? parsed : parsed ? [parsed] : [];
    const failed = prepared.filter((item) => item.wasDirty && !item.saved);
    if (failed.length > 0) {
      throw new Error(`无法保存 ${failed.length} 个已打开的 Office 文档，事务已停止: ${failed.map((item) => item.filePath).join(", ")}`);
    }
    return prepared;
  }

  async restoreTransactionFiles(files: Array<{
    filePath: string;
    existed: boolean;
    snapshotPath?: string;
  }>): Promise<Array<{ app: OfficeActionApp; filePath: string; instanceId?: string; reopened: boolean }>> {
    const output = await executePowerShell(`
${officeInstanceDiscoveryScript()}
${psVar("_restoreFilesJson", JSON.stringify(files))}
$restoreFiles = @(ConvertFrom-Json $_restoreFilesJson)
$restorePaths = @($restoreFiles | ForEach-Object { [IO.Path]::GetFullPath([string]$_.filePath) })
$sessions = @()
foreach ($handle in @(Get-AllOfficeDocumentHandles '')) {
  $fullName = try { [IO.Path]::GetFullPath([string]$handle.document.FullName) } catch { '' }
  if (-not $fullName -or $fullName -notin $restorePaths) { continue }
  $wasDirty = try { -not [bool]$handle.document.Saved } catch { $false }
  if ($wasDirty) { $handle.document.Save() }
  $metadata = Get-OfficeDocumentMetadata $handle
  $sessions += [pscustomobject]@{
    app = [string]$handle.app
    application = $handle.application
    filePath = $fullName
    instanceId = [string]$handle.instanceId
    active = [bool]$metadata.active
    readOnly = [bool]$metadata.readOnly
  }
  $handle.document.Close($false)
}

$published = @()
$stagedFiles = @()
$committedFiles = @()
try {
  foreach ($file in $restoreFiles) {
    $destination = [IO.Path]::GetFullPath([string]$file.filePath)
    $directory = [IO.Path]::GetDirectoryName($destination)
    if (-not [IO.Directory]::Exists($directory)) { [void][IO.Directory]::CreateDirectory($directory) }
    $stagedPath = ''
    if ([bool]$file.existed) {
      $source = [IO.Path]::GetFullPath([string]$file.snapshotPath)
      if (-not [IO.File]::Exists($source)) { throw 'Office 事务快照不存在: ' + $source }
      $stagedPath = $destination + '.' + [Guid]::NewGuid().ToString('N') + '.transaction.stage'
      [IO.File]::Copy($source, $stagedPath, $true)
    }
    $stagedFiles += [pscustomobject]@{ destination = $destination; stagedPath = $stagedPath; rollbackPath = ''; committed = $false }
  }
  foreach ($entry in $stagedFiles) {
    if ([IO.File]::Exists([string]$entry.destination)) {
      $entry.rollbackPath = [string]$entry.destination + '.' + [Guid]::NewGuid().ToString('N') + '.transaction.rollback'
      [IO.File]::Move([string]$entry.destination, [string]$entry.rollbackPath)
    }
    $entry.committed = $true
    $committedFiles += $entry
    if ($entry.stagedPath) {
      [IO.File]::Move([string]$entry.stagedPath, [string]$entry.destination)
      $entry.stagedPath = ''
    }
  }
  foreach ($entry in $stagedFiles) { if ($entry.rollbackPath -and [IO.File]::Exists([string]$entry.rollbackPath)) { [IO.File]::Delete([string]$entry.rollbackPath); $entry.rollbackPath = '' } }
} catch {
  for ($index = $committedFiles.Count - 1; $index -ge 0; $index--) {
    $entry = $committedFiles[$index]
    if ([IO.File]::Exists([string]$entry.destination)) { [IO.File]::Delete([string]$entry.destination) }
    if ($entry.rollbackPath -and [IO.File]::Exists([string]$entry.rollbackPath)) { [IO.File]::Move([string]$entry.rollbackPath, [string]$entry.destination); $entry.rollbackPath = '' }
  }
  throw
} finally {
  foreach ($entry in $stagedFiles) {
    if ($entry.stagedPath -and [IO.File]::Exists([string]$entry.stagedPath)) { [IO.File]::Delete([string]$entry.stagedPath) }
    if ($entry.rollbackPath -and [IO.File]::Exists([string]$entry.rollbackPath)) { [IO.File]::Delete([string]$entry.rollbackPath) }
  }
  foreach ($session in $sessions) {
    $reopened = $false
    if ([IO.File]::Exists([string]$session.filePath)) {
      $document = $null
      try {
        switch ([string]$session.app) {
          'excel' { $document = $session.application.Workbooks.Open([string]$session.filePath, 0, [bool]$session.readOnly) }
          'word' { $document = $session.application.Documents.Open([string]$session.filePath, $false, [bool]$session.readOnly) }
          default { $document = $session.application.Presentations.Open([string]$session.filePath, [bool]$session.readOnly, $false, -1) }
        }
        $reopened = $null -ne $document
        if ($reopened -and [bool]$session.active) { $document.Activate() }
      } catch {}
    }
    $published += [pscustomobject]@{ app = [string]$session.app; filePath = [string]$session.filePath; instanceId = [string]$session.instanceId; reopened = $reopened }
  }
}
$published | ConvertTo-Json -Depth 5 -Compress
`);
    const parsed = safeJsonParse<
      { app: OfficeActionApp; filePath: string; instanceId?: string; reopened: boolean }
      | Array<{ app: OfficeActionApp; filePath: string; instanceId?: string; reopened: boolean }>
    >(output || "[]", "powershell", "恢复已打开的 Office 事务文件");
    const restored = Array.isArray(parsed) ? parsed : parsed ? [parsed] : [];
    const failed = restored.filter((item) => !item.reopened && files.some((file) => file.existed && pathKey(file.filePath) === pathKey(item.filePath)));
    if (failed.length > 0) throw new Error(`事务文件已恢复，但 ${failed.length} 个 Office 文档无法重新打开: ${failed.map((item) => item.filePath).join(", ")}`);
    return restored;
  }
}

function pathKey(filePath: string): string {
  return path.resolve(filePath).toLowerCase();
}

export function officeInstanceDiscoveryScript(): string {
  return String.raw`
if (-not ('WenggeRotEntry' -as [type])) {
  Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

public sealed class WenggeRotEntry {
  public string DisplayName { get; set; }
  public object Value { get; set; }
}

public static class WenggeRotDiscovery {
  [DllImport("ole32.dll")]
  private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable runningObjectTable);
  [DllImport("ole32.dll")]
  private static extern int CreateBindCtx(int reserved, out IBindCtx bindContext);

  public static WenggeRotEntry[] Enumerate() {
    var result = new List<WenggeRotEntry>();
    IRunningObjectTable table = null;
    IBindCtx context = null;
    IEnumMoniker iterator = null;
    try {
      if (GetRunningObjectTable(0, out table) != 0 || table == null) return result.ToArray();
      if (CreateBindCtx(0, out context) != 0 || context == null) return result.ToArray();
      table.EnumRunning(out iterator);
      if (iterator == null) return result.ToArray();
      iterator.Reset();
      var monikers = new IMoniker[1];
      while (iterator.Next(1, monikers, IntPtr.Zero) == 0) {
        var moniker = monikers[0];
        try {
          string displayName;
          moniker.GetDisplayName(context, null, out displayName);
          object value;
          table.GetObject(moniker, out value);
          if (value != null) result.Add(new WenggeRotEntry { DisplayName = displayName ?? "", Value = value });
        } catch { }
        finally {
          if (moniker != null && Marshal.IsComObject(moniker)) Marshal.ReleaseComObject(moniker);
        }
      }
    } finally {
      if (iterator != null && Marshal.IsComObject(iterator)) Marshal.ReleaseComObject(iterator);
      if (context != null && Marshal.IsComObject(context)) Marshal.ReleaseComObject(context);
      if (table != null && Marshal.IsComObject(table)) Marshal.ReleaseComObject(table);
    }
    return result.ToArray();
  }
}

public static class WenggeOfficeWindowProcess {
  [DllImport("user32.dll")]
  public static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
}
'@
}

function Get-OfficeDocumentKind($document) {
  $saved = $null
  try { $saved = $document.Saved } catch {}
  if ($null -eq $saved) { return '' }
  $fullName = try { [string]$document.FullName } catch { '' }
  $extension = try { [IO.Path]::GetExtension($fullName).ToLowerInvariant() } catch { '' }
  if ($extension -in @('.xlsx', '.xlsm', '.xlsb', '.xls', '.xltx', '.xltm', '.et')) { return 'excel' }
  if ($extension -in @('.docx', '.docm', '.doc', '.dotx', '.dotm', '.wps')) { return 'word' }
  if ($extension -in @('.pptx', '.pptm', '.ppt', '.potx', '.potm', '.dps')) { return 'presentation' }
  $worksheets = $null
  try { $worksheets = $document.Worksheets } catch {}
  if ($null -ne $worksheets) { return 'excel' }
  $content = $null; $bookmarks = $null
  try { $content = $document.Content; $bookmarks = $document.Bookmarks } catch {}
  if ($null -ne $content -and $null -ne $bookmarks) { return 'word' }
  $slides = $null
  try { $slides = $document.Slides } catch {}
  if ($null -ne $slides) { return 'presentation' }
  return ''
}

function Get-OfficeApplicationInfo($application, $document, [string]$appKind, [string]$candidateProgId, [string]$rotName) {
  $hwnd = [int64]0
  try { $hwnd = [int64]$application.Hwnd } catch { try { $hwnd = [int64]$application.HWND } catch {} }
  if ($hwnd -eq 0) {
    switch ($appKind) {
      'excel' { try { $hwnd = [int64]$document.Windows.Item(1).Hwnd } catch {} }
      'word' {
        try { $hwnd = [int64]$document.ActiveWindow.Hwnd } catch {}
        if ($hwnd -eq 0) { try { $hwnd = [int64]$document.Windows.Item(1).Hwnd } catch {} }
      }
      default {
        try { $hwnd = [int64]$document.Windows.Item(1).HWND } catch {}
        if ($hwnd -eq 0) { try { $hwnd = [int64]$application.ActiveWindow.HWND } catch {} }
      }
    }
  }
  $processId = [uint32]0
  if ($hwnd -ne 0) { try { [void][WenggeOfficeWindowProcess]::GetWindowThreadProcessId([IntPtr]$hwnd, [ref]$processId) } catch {} }
  $processName = ''
  if ($processId -gt 0) { try { $processName = [string](Get-Process -Id $processId -ErrorAction Stop).ProcessName } catch {} }
  $applicationName = try { [string]$application.Name } catch { '' }
  $officeHost = if ($processName -match '^(EXCEL|WINWORD|POWERPNT)$' -or $candidateProgId -match '^(Excel|Word|PowerPoint)\.' -or $applicationName -match '(?i)Microsoft|Excel|Word|PowerPoint') {
    'microsoft-office'
  } elseif ($processName -match '^(wps|et|wpp|kso)' -or $candidateProgId -match '(?i)wps|ket' -or $applicationName -match '(?i)WPS|Kingsoft') {
    'wps'
  } else { 'unknown' }
  $progId = $candidateProgId
  if (-not $progId) {
    if ($officeHost -eq 'microsoft-office') {
      $progId = @{ excel = 'Excel.Application'; word = 'Word.Application'; presentation = 'PowerPoint.Application' }[$appKind]
    } elseif ($officeHost -eq 'wps') {
      $progId = @{ excel = 'Ket.Application'; word = 'Wps.Application'; presentation = 'Wpp.Application' }[$appKind]
    } else { $progId = 'ROT' }
  }
  $instanceId = if ($processId -gt 0 -or $hwnd -ne 0) {
    $appKind + ':' + [string]$processId + ':' + [string]$hwnd
  } else {
    $appKind + ':rot:' + [Uri]::EscapeDataString($rotName)
  }
  return [pscustomobject]@{ processId = [int]$processId; hwnd = $hwnd; host = $officeHost; progId = $progId; instanceId = $instanceId }
}

function New-OfficeDocumentHandle($document, [string]$appKind, [string]$candidateProgId, [string]$rotName) {
  if ($null -eq $document -or -not $appKind) { return $null }
  $application = try { $document.Application } catch { $null }
  if ($null -eq $application) { return $null }
  try {
    $documentName = [string]$document.Name
    $applicationName = [string]$application.Name
    if (-not $documentName -or -not $applicationName) { return $null }
  } catch { return $null }
  $applicationInfo = Get-OfficeApplicationInfo $application $document $appKind $candidateProgId $rotName
  return [pscustomobject]@{
    app = $appKind
    document = $document
    application = $application
    instanceId = [string]$applicationInfo.instanceId
    processId = [int]$applicationInfo.processId
    hwnd = [int64]$applicationInfo.hwnd
    host = [string]$applicationInfo.host
    progId = [string]$applicationInfo.progId
    rotName = $rotName
  }
}

function Get-OfficeDocumentMetadata($handle) {
  $document = $handle.document
  $fullName = try { [string]$document.FullName } catch { '' }
  $activeObject = switch ([string]$handle.app) {
    'excel' { try { $handle.application.ActiveWorkbook } catch { $null } }
    'word' { try { $handle.application.ActiveDocument } catch { $null } }
    default { try { $handle.application.ActivePresentation } catch { $null } }
  }
  $activeFullName = try { [string]$activeObject.FullName } catch { '' }
  return [pscustomobject]@{
    app = [string]$handle.app
    name = $(try { [string]$document.Name } catch { '' })
    fullName = $fullName
    index = $(try { [int]$document.Index } catch { 0 })
    active = [bool]($activeFullName -and $fullName -and $activeFullName -ieq $fullName)
    progId = [string]$handle.progId
    host = [string]$handle.host
    instanceId = [string]$handle.instanceId
    processId = [int]$handle.processId
    hwnd = [int64]$handle.hwnd
    readOnly = $(try { [bool]$document.ReadOnly } catch { $false })
    saved = $(try { [bool]$document.Saved } catch { $true })
  }
}

function Get-AllOfficeDocumentHandles([string]$appFilter) {
  $handles = @()
  $seen = @{}
  $seenPaths = @{}
  foreach ($entry in @([WenggeRotDiscovery]::Enumerate())) {
    $document = $entry.Value
    $appKind = Get-OfficeDocumentKind $document
    if (-not $appKind -or ($appFilter -and $appFilter -ne $appKind)) { continue }
    $handle = New-OfficeDocumentHandle $document $appKind '' ([string]$entry.DisplayName)
    if ($null -eq $handle) { continue }
    $fullName = try { [IO.Path]::GetFullPath([string]$document.FullName).ToLowerInvariant() } catch { [string]$document.FullName }
    $pathKey = ($appKind + '|' + $fullName).ToLowerInvariant()
    if ($fullName -and $seenPaths.ContainsKey($pathKey)) { continue }
    $key = ([string]$handle.instanceId + '|' + $appKind + '|' + $fullName + '|' + $(try { [int]$document.Index } catch { 0 })).ToLowerInvariant()
    if (-not $seen.ContainsKey($key)) { $seen[$key] = $true; if ($fullName) { $seenPaths[$pathKey] = $true }; $handles += $handle }
  }

  $configs = @(
    [pscustomobject]@{ app = 'excel'; progIds = @('Excel.Application', 'Ket.Application'); collection = 'Workbooks' },
    [pscustomobject]@{ app = 'word'; progIds = @('Word.Application', 'Kwps.Application', 'Wps.Application'); collection = 'Documents' },
    [pscustomobject]@{ app = 'presentation'; progIds = @('PowerPoint.Application', 'Wpp.Application', 'Kwpp.Application'); collection = 'Presentations' }
  )
  foreach ($config in $configs) {
    if ($appFilter -and $appFilter -ne $config.app) { continue }
    foreach ($candidateProgId in $config.progIds) {
      $application = $null
      try { $application = [Runtime.InteropServices.Marshal]::GetActiveObject([string]$candidateProgId) } catch {}
      if ($null -eq $application) { continue }
      foreach ($document in @($application.($config.collection))) {
        $handle = New-OfficeDocumentHandle $document ([string]$config.app) ([string]$candidateProgId) ''
        if ($null -eq $handle) { continue }
        $fullName = try { [IO.Path]::GetFullPath([string]$document.FullName).ToLowerInvariant() } catch { [string]$document.FullName }
        $pathKey = ([string]$config.app + '|' + $fullName).ToLowerInvariant()
        if ($fullName -and $seenPaths.ContainsKey($pathKey)) { continue }
        $key = ([string]$handle.instanceId + '|' + [string]$config.app + '|' + $fullName + '|' + $(try { [int]$document.Index } catch { 0 })).ToLowerInvariant()
        if (-not $seen.ContainsKey($key)) { $seen[$key] = $true; if ($fullName) { $seenPaths[$pathKey] = $true }; $handles += $handle }
      }
    }
  }
  return @($handles | Sort-Object app, processId, hwnd, @{ Expression = { try { [int]$_.document.Index } catch { 0 } } })
}

function Resolve-OfficeDocumentHandle([string]$appKind, [string]$filePath, [string]$instanceId, [string]$name, [int]$index) {
  $candidates = @(Get-AllOfficeDocumentHandles $appKind)
  if ($instanceId) { $candidates = @($candidates | Where-Object { $_.instanceId -eq $instanceId }) }
  if ($filePath) {
    $wanted = [IO.Path]::GetFullPath($filePath)
    $candidates = @($candidates | Where-Object { try { [IO.Path]::GetFullPath([string]$_.document.FullName) -ieq $wanted } catch { $false } })
  } elseif ($name) {
    $candidates = @($candidates | Where-Object { try { [string]$_.document.Name -ieq $name } catch { $false } })
  } elseif ($index -gt 0) {
    $candidates = @($candidates | Where-Object { try { [int]$_.document.Index -eq $index } catch { $false } })
  } elseif (-not $instanceId) {
    throw '需要 instanceId、filePath、name 或 index 之一'
  }
  if ($candidates.Count -eq 0) { throw '找不到指定实例和完整路径的 Office 文档窗口' }
  if ($candidates.Count -gt 1) { throw '找到多个 Office 文档候选，请传 office.documents.list 返回的 instanceId 和完整路径' }
  return $candidates[0]
}
`;
}

function officeDocumentResolverScript(app: OfficeActionApp, filePath: string, instanceId?: string): string {
  return `
${officeInstanceDiscoveryScript()}
${psVar("_filePath", filePath)}
${psVar("_instanceId", instanceId || "")}
$handle = Resolve-OfficeDocumentHandle '${app}' $_filePath $_instanceId '' 0
$app = $handle.application
$target = $handle.document
`;
}

function listObjectsScript(app: OfficeActionApp): string {
  if (app === "excel") return `
$activeSheetName = try { [string]$target.ActiveSheet.Name } catch { '' }
foreach ($sheet in $target.Worksheets) {
  $sheetName = [string]$sheet.Name; $sheetPart = Encode-LocatorPart $sheetName
  Add-OfficeObject 'sheet' $sheetName ('sheet:' + $sheetPart) '' ([int]$sheet.Index) '' ($sheetName -eq $activeSheetName)
  try { $used = $sheet.UsedRange; $address = [string]$used.Address($false, $false); Add-OfficeObject 'range' $address ('range:' + $sheetPart + '/' + (Encode-LocatorPart $address)) $sheetName 1 ('rows=' + $used.Rows.Count + ';columns=' + $used.Columns.Count) } catch {}
  try { foreach ($table in $sheet.ListObjects) { Add-OfficeObject 'table' ([string]$table.Name) ('table:' + $sheetPart + '/' + (Encode-LocatorPart ([string]$table.Name))) $sheetName ([int]$table.Index) ([string]$table.Range.Address($false, $false)) } } catch {}
  try { foreach ($chart in $sheet.ChartObjects()) { Add-OfficeObject 'chart' ([string]$chart.Name) ('chart:' + $sheetPart + '/' + (Encode-LocatorPart ([string]$chart.Name))) $sheetName ([int]$chart.Index) } } catch {}
  try { foreach ($shape in $sheet.Shapes) { Add-OfficeObject 'shape' ([string]$shape.Name) ('shape:' + $sheetPart + '/' + (Encode-LocatorPart ([string]$shape.Name))) $sheetName ([int]$shape.Id) } } catch {}
  try { foreach ($pivot in $sheet.PivotTables()) { Add-OfficeObject 'pivotTable' ([string]$pivot.Name) ('pivotTable:' + $sheetPart + '/' + (Encode-LocatorPart ([string]$pivot.Name))) $sheetName 0 ([string]$pivot.TableRange2.Address($false, $false)) } } catch {}
}
try {
  $activeCell = $app.ActiveCell
  if ($null -ne $activeCell -and [string]$activeCell.Parent.Parent.FullName -ieq [string]$target.FullName) {
    $sheetName = [string]$activeCell.Worksheet.Name; $address = [string]$activeCell.Address($false, $false)
    Add-OfficeObject 'cell' $address ('cell:' + (Encode-LocatorPart $sheetName) + '/' + (Encode-LocatorPart $address)) $sheetName 0 '' $true
  }
} catch {}
try {
  $selection = $app.Selection
  if ($null -ne $selection -and [string]$selection.Parent.Parent.FullName -ieq [string]$target.FullName) {
    $sheetName = [string]$selection.Worksheet.Name; $address = [string]$selection.Address($false, $false)
    Add-OfficeObject 'range' $address ('range:' + (Encode-LocatorPart $sheetName) + '/' + (Encode-LocatorPart $address)) $sheetName 0 '当前选区' $true
  }
} catch {}
try { foreach ($name in $target.Names) { Add-OfficeObject 'name' ([string]$name.Name) ('name:' + (Encode-LocatorPart ([string]$name.Name))) '' ([int]$name.Index) ([string]$name.RefersTo) } } catch {}
try { foreach ($query in $target.Queries) { Add-OfficeObject 'query' ([string]$query.Name) ('query:' + (Encode-LocatorPart ([string]$query.Name))) '' 0 ([string]$query.Formula) } } catch {}
try { foreach ($connection in $target.Connections) { Add-OfficeObject 'connection' ([string]$connection.Name) ('connection:' + (Encode-LocatorPart ([string]$connection.Name))) '' 0 ([string]$connection.Description) } } catch {}
try {
  foreach ($cache in $target.SlicerCaches) {
    foreach ($slicer in $cache.Slicers) {
      Add-OfficeObject 'slicer' ([string]$slicer.Name) ('slicer:' + (Encode-LocatorPart ([string]$cache.Name)) + '/' + (Encode-LocatorPart ([string]$slicer.Name))) ([string]$cache.Name) 0 ([string]$slicer.Caption)
    }
  }
} catch {}
`;
  if (app === "word") return `
$pageCount = try { [int]$target.ComputeStatistics(2) } catch { 0 }
if ($pageCount -gt 0) { foreach ($pageIndex in 1..$pageCount) { Add-OfficeObject 'page' ('第 ' + $pageIndex + ' 页') ('page:' + $pageIndex) '' $pageIndex } }
for ($index = 1; $index -le $target.Sections.Count; $index++) { Add-OfficeObject 'section' ('第 ' + $index + ' 节') ('section:' + $index) '' $index }
foreach ($paragraph in $target.Paragraphs) {
  $start = [int]$paragraph.Range.Start
  $text = ([string]$paragraph.Range.Text).Trim(); if ($text.Length -gt 120) { $text = $text.Substring(0, 120) }
  Add-OfficeObject 'paragraph' $(if ($text) { $text } else { '空段落 ' + $start }) ('paragraph:' + $start) '' $start $text
  $outlineLevel = try { [int]$paragraph.OutlineLevel } catch { 10 }
  if ($outlineLevel -ge 1 -and $outlineLevel -le 9) { Add-OfficeObject 'heading' $(if ($text) { $text } else { '标题 ' + $start }) ('heading:' + $start) '' $start ('level=' + $outlineLevel) }
  $styleName = try { [string]$paragraph.Range.Style.NameLocal } catch { try { [string]$paragraph.Range.Style } catch { '' } }
  if ($styleName -match '(?i)caption|题注') { Add-OfficeObject 'caption' $(if ($text) { $text } else { '题注 ' + $start }) ('caption:' + $start) '' $start $styleName }
}
for ($index = 1; $index -le $target.Tables.Count; $index++) { $table = $target.Tables.Item($index); Add-OfficeObject 'table' ('表格 ' + $index) ('table:' + $index) '' $index ('rows=' + $table.Rows.Count + ';columns=' + $table.Columns.Count) }
foreach ($bookmark in $target.Bookmarks) { Add-OfficeObject 'bookmark' ([string]$bookmark.Name) ('bookmark:' + (Encode-LocatorPart ([string]$bookmark.Name))) '' ([int]$bookmark.Range.Start) }
foreach ($control in $target.ContentControls) { Add-OfficeObject 'contentControl' ([string]$(if ($control.Title) { $control.Title } elseif ($control.Tag) { $control.Tag } else { $control.ID })) ('contentControl:' + [string]$control.ID) '' ([int64]$control.ID) ([string]$control.Tag) }
for ($index = 1; $index -le $target.InlineShapes.Count; $index++) { Add-OfficeObject 'inlineShape' ('嵌入对象 ' + $index) ('inlineShape:' + $index) '' $index }
foreach ($shape in $target.Shapes) { Add-OfficeObject 'shape' ([string]$shape.Name) ('shape:' + (Encode-LocatorPart ([string]$shape.Name))) '' ([int]$shape.ID) }
for ($index = 1; $index -le $target.Comments.Count; $index++) { $item = $target.Comments.Item($index); $text = ([string]$item.Range.Text).Trim(); Add-OfficeObject 'comment' $(if ($text) { $text } else { '批注 ' + $index }) ('comment:' + $index) '' $index ([string]$item.Author) }
for ($index = 1; $index -le $target.Revisions.Count; $index++) { $item = $target.Revisions.Item($index); Add-OfficeObject 'revision' ('修订 ' + $index) ('revision:' + $index) '' $index ('type=' + [int]$item.Type + ';author=' + [string]$item.Author) }
for ($index = 1; $index -le $target.Footnotes.Count; $index++) { $item = $target.Footnotes.Item($index); Add-OfficeObject 'footnote' ('脚注 ' + $index) ('footnote:' + $index) '' $index (([string]$item.Range.Text).Trim()) }
for ($index = 1; $index -le $target.Endnotes.Count; $index++) { $item = $target.Endnotes.Item($index); Add-OfficeObject 'endnote' ('尾注 ' + $index) ('endnote:' + $index) '' $index (([string]$item.Range.Text).Trim()) }
`;
  return `
$activeSlideIndex = try { [int]$app.ActiveWindow.View.Slide.SlideIndex } catch { 0 }
function Add-PresentationShapeObject($shape, [int]$slideId, [string]$parentLocator, [string]$path) {
  $shapePart = Encode-LocatorPart ([string]$shape.Name)
  $shapePath = $(if ($path) { $path + '/' + $shapePart } else { $shapePart })
  $detail = try { [string]$shape.TextFrame.TextRange.Text } catch { '' }; if ($detail.Length -gt 120) { $detail = $detail.Substring(0, 120) }
  $locator = 'shape:' + $slideId + '/' + $shapePath
  Add-OfficeObject 'shape' ([string]$shape.Name) $locator $parentLocator ([int]$shape.Id) $detail
  try { if ($shape.HasChart) { Add-OfficeObject 'chart' ([string]$shape.Name) ('chart:' + $slideId + '/' + $shapePath) $parentLocator ([int]$shape.Id) } } catch {}
  try { if ($shape.HasTable) { Add-OfficeObject 'table' ([string]$shape.Name) ('table:' + $slideId + '/' + $shapePath) $parentLocator ([int]$shape.Id) } } catch {}
  try {
    if ([int]$shape.Type -eq 6) {
      for ($childIndex = 1; $childIndex -le $shape.GroupItems.Count; $childIndex++) {
        Add-PresentationShapeObject $shape.GroupItems.Item($childIndex) $slideId $locator $shapePath
      }
    }
  } catch {}
}
for ($designIndex = 1; $designIndex -le $target.Designs.Count; $designIndex++) {
  $design = $target.Designs.Item($designIndex)
  Add-OfficeObject 'master' ([string]$design.Name) ('master:' + $designIndex) '' $designIndex ([string]$design.SlideMaster.Name)
  for ($layoutIndex = 1; $layoutIndex -le $design.SlideMaster.CustomLayouts.Count; $layoutIndex++) {
    $layout = $design.SlideMaster.CustomLayouts.Item($layoutIndex)
    Add-OfficeObject 'layout' ([string]$layout.Name) ('layout:' + $designIndex + '/' + $layoutIndex) ('master:' + $designIndex) $layoutIndex
  }
}
foreach ($slide in $target.Slides) {
  $slideIndex = [int]$slide.SlideIndex; $slideId = [int]$slide.SlideID; $title = try { [string]$slide.Shapes.Title.TextFrame.TextRange.Text } catch { '' }
  Add-OfficeObject 'slide' $(if ($title) { $title } else { '幻灯片 ' + $slideIndex }) ('slide:' + $slideId) '' $slideIndex ('slideId=' + $slideId + ';title=' + $title) ($slideIndex -eq $activeSlideIndex)
  Add-OfficeObject 'notesPage' ('幻灯片 ' + $slideIndex + ' 备注页') ('notesPage:' + $slideId) ('slide:' + $slideId) $slideIndex
  foreach ($shape in $slide.Shapes) { Add-PresentationShapeObject $shape $slideId ('slide:' + $slideId) '' }
}
`;
}

function activateObjectScript(app: OfficeActionApp): string {
  if (app === "excel") return `
if ($_locator -match '^(sheet|cell|range|table|chart|shape|pivotTable):(.+)$') {
  $selectedKind = $Matches[1]; $value = $Matches[2]
  if ($selectedKind -eq 'sheet') { $sheetName = Decode-LocatorPart $value; $sheet = $target.Worksheets.Item($sheetName); $sheet.Activate(); $selectedName = $sheetName; $selectedIndex = [int]$sheet.Index }
  else {
    $parts = @($value -split '/', 2); if ($parts.Count -ne 2) { throw 'Excel 对象 locator 格式无效' }
    $sheetName = Decode-LocatorPart $parts[0]; $objectName = Decode-LocatorPart $parts[1]; $sheet = $target.Worksheets.Item($sheetName); $sheet.Activate(); $selectedParent = $sheetName
    switch ($selectedKind) {
      'cell' { $object = $sheet.Range($objectName); $object.Select(); $selectedName = $objectName }
      'range' { $object = $sheet.Range($objectName); $object.Select(); $selectedName = $objectName }
      'table' { $object = $sheet.ListObjects.Item($objectName); $object.Range.Select(); $selectedName = [string]$object.Name; $selectedIndex = [int]$object.Index }
      'chart' { $object = $sheet.ChartObjects($objectName); $object.Activate(); $selectedName = [string]$object.Name; $selectedIndex = [int]$object.Index }
      'shape' { $object = $sheet.Shapes.Item($objectName); $object.Select(); $selectedName = [string]$object.Name; $selectedIndex = [int]$object.Id }
      'pivotTable' { $object = $sheet.PivotTables($objectName); $object.TableRange2.Select(); $selectedName = [string]$object.Name }
    }
  }
} elseif ($_locator -match '^name:(.+)$') { $selectedKind = 'name'; $selectedName = Decode-LocatorPart $Matches[1]; $object = $target.Names.Item($selectedName); $object.RefersToRange.Select(); $selectedIndex = [int]$object.Index }
elseif ($_locator -match '^slicer:([^/]+)/(.+)$') { $selectedKind = 'slicer'; $cacheName = Decode-LocatorPart $Matches[1]; $selectedName = Decode-LocatorPart $Matches[2]; $object = $target.SlicerCaches.Item($cacheName).Slicers.Item($selectedName); $object.Shape.Select() }
elseif ($_locator -match '^(query|connection):(.+)$') {
  $selectedKind = $Matches[1]; $selectedName = Decode-LocatorPart $Matches[2]; $target.Activate()
  foreach ($candidateSheet in $target.Worksheets) {
    foreach ($listObject in $candidateSheet.ListObjects) {
      $connectionName = try { [string]$listObject.QueryTable.WorkbookConnection.Name } catch { '' }
      if ($selectedKind -eq 'connection' -and $connectionName -ieq $selectedName) { $candidateSheet.Activate(); $listObject.Range.Select(); break }
      if ($selectedKind -eq 'query' -and ($connectionName -ieq ('Query - ' + $selectedName) -or $connectionName -ieq $selectedName)) { $candidateSheet.Activate(); $listObject.Range.Select(); break }
    }
  }
}
`;
  if (app === "word") return `
if ($_locator -match '^page:([0-9]+)$') { $selectedKind = 'page'; $selectedIndex = [int]$Matches[1]; $selectedName = '第 ' + $selectedIndex + ' 页'; [void]$app.Selection.GoTo(1, 1, $selectedIndex) }
elseif ($_locator -match '^section:([0-9]+)$') { $selectedKind = 'section'; $selectedIndex = [int]$Matches[1]; $selectedName = '第 ' + $selectedIndex + ' 节'; $target.Sections.Item($selectedIndex).Range.Select() }
elseif ($_locator -match '^(paragraph|heading|caption):([0-9]+)$') { $selectedKind = $Matches[1]; $rangeStart = [int]$Matches[2]; foreach ($paragraph in $target.Paragraphs) { if ([int]$paragraph.Range.Start -eq $rangeStart) { $paragraph.Range.Select(); $selectedName = ([string]$paragraph.Range.Text).Trim(); $selectedIndex = $rangeStart; break } } }
elseif ($_locator -match '^table:([0-9]+)$') { $selectedKind = 'table'; $selectedIndex = [int]$Matches[1]; $selectedName = '表格 ' + $selectedIndex; $target.Tables.Item($selectedIndex).Range.Select() }
elseif ($_locator -match '^bookmark:(.+)$') { $selectedKind = 'bookmark'; $selectedName = Decode-LocatorPart $Matches[1]; $target.Bookmarks.Item($selectedName).Range.Select() }
elseif ($_locator -match '^contentControl:([0-9]+)$') { $selectedKind = 'contentControl'; $selectedIndex = [int]$Matches[1]; foreach ($control in $target.ContentControls) { if ([int64]$control.ID -eq [int64]$Matches[1]) { $selectedName = $(if ($control.Title) { [string]$control.Title } else { [string]$control.ID }); $control.Range.Select(); break } } }
elseif ($_locator -match '^inlineShape:([0-9]+)$') { $selectedKind = 'inlineShape'; $selectedIndex = [int]$Matches[1]; $selectedName = '嵌入对象 ' + $selectedIndex; $target.InlineShapes.Item($selectedIndex).Range.Select() }
elseif ($_locator -match '^shape:(.+)$') { $selectedKind = 'shape'; $selectedName = Decode-LocatorPart $Matches[1]; $object = $target.Shapes.Item($selectedName); $object.Select(); $selectedIndex = [int]$object.ID }
elseif ($_locator -match '^comment:([0-9]+)$') { $selectedKind = 'comment'; $selectedIndex = [int]$Matches[1]; $object = $target.Comments.Item($selectedIndex); $object.Scope.Select(); $selectedName = ([string]$object.Range.Text).Trim() }
elseif ($_locator -match '^revision:([0-9]+)$') { $selectedKind = 'revision'; $selectedIndex = [int]$Matches[1]; $object = $target.Revisions.Item($selectedIndex); $object.Range.Select(); $selectedName = '修订 ' + $selectedIndex }
elseif ($_locator -match '^footnote:([0-9]+)$') { $selectedKind = 'footnote'; $selectedIndex = [int]$Matches[1]; $object = $target.Footnotes.Item($selectedIndex); $object.Reference.Select(); $selectedName = '脚注 ' + $selectedIndex }
elseif ($_locator -match '^endnote:([0-9]+)$') { $selectedKind = 'endnote'; $selectedIndex = [int]$Matches[1]; $object = $target.Endnotes.Item($selectedIndex); $object.Reference.Select(); $selectedName = '尾注 ' + $selectedIndex }
`;
  return `
function Get-SlideById([int]$slideId) { foreach ($candidateSlide in $target.Slides) { if ([int]$candidateSlide.SlideID -eq $slideId) { return $candidateSlide } }; return $null }
function Get-ShapeByPath($slide, [string[]]$parts) {
  if ($parts.Count -eq 0) { return $null }
  $shape = $slide.Shapes.Item((Decode-LocatorPart $parts[0]))
  for ($pathIndex = 1; $pathIndex -lt $parts.Count; $pathIndex++) { $shape = $shape.GroupItems.Item((Decode-LocatorPart $parts[$pathIndex])) }
  return $shape
}
if ($_locator -match '^slide:([0-9]+)$') { $selectedKind = 'slide'; $slide = Get-SlideById ([int]$Matches[1]); if ($null -eq $slide) { throw '找不到指定 SlideID' }; $slide.Select(); $selectedIndex = [int]$slide.SlideIndex; $selectedName = '幻灯片 ' + $selectedIndex }
elseif ($_locator -match '^notesPage:([0-9]+)$') { $selectedKind = 'notesPage'; $slide = Get-SlideById ([int]$Matches[1]); if ($null -eq $slide) { throw '找不到指定 SlideID' }; $app.ActiveWindow.View.GotoSlide([int]$slide.SlideIndex); $app.ActiveWindow.ViewType = 3; $selectedIndex = [int]$slide.SlideIndex; $selectedName = '幻灯片 ' + $selectedIndex + ' 备注页' }
elseif ($_locator -match '^master:([0-9]+)$') { $selectedKind = 'master'; $selectedIndex = [int]$Matches[1]; $object = $target.Designs.Item($selectedIndex).SlideMaster; $app.ActiveWindow.ViewType = 2; try { $object.Select() } catch {}; $selectedName = [string]$object.Name }
elseif ($_locator -match '^layout:([0-9]+)/([0-9]+)$') { $selectedKind = 'layout'; $designIndex = [int]$Matches[1]; $selectedIndex = [int]$Matches[2]; $object = $target.Designs.Item($designIndex).SlideMaster.CustomLayouts.Item($selectedIndex); $app.ActiveWindow.ViewType = 2; try { $object.Select() } catch { try { $target.Designs.Item($designIndex).SlideMaster.Select() } catch {} }; $selectedName = [string]$object.Name; $selectedParent = 'master:' + $designIndex }
elseif ($_locator -match '^(shape|chart|table):([0-9]+)/(.+)$') { $selectedKind = $Matches[1]; $slideId = [int]$Matches[2]; $parts = @($Matches[3] -split '/'); $slide = Get-SlideById $slideId; if ($null -eq $slide) { throw '找不到指定 SlideID' }; $slide.Select(); $object = Get-ShapeByPath $slide $parts; if ($null -eq $object) { throw '找不到指定形状路径' }; $object.Select(); $selectedName = [string]$object.Name; $selectedParent = 'slide:' + $slideId; $selectedIndex = [int]$object.Id }
`;
}
