const WORD_CONTENT_CONTROL_OPERATIONS = new Set([
  "inspectContentControls",
  "populateContentControls",
  "manageContentControls",
]);

export function buildWordContentControlOperationScript(operation: string): string | undefined {
  if (!WORD_CONTENT_CONTROL_OPERATIONS.has(operation)) return undefined;
  return String.raw`
  function Get-ContentControlTypeName([int]$type) {
    switch ($type) {
      0 { 'richText' }
      1 { 'text' }
      2 { 'picture' }
      3 { 'comboBox' }
      4 { 'dropDownList' }
      5 { 'buildingBlockGallery' }
      6 { 'date' }
      7 { 'group' }
      8 { 'checkBox' }
      9 { 'repeatingSection' }
      default { 'unknown' }
    }
  }
  function Resolve-ContentControlType([string]$type) {
    switch ($type) {
      'richText' { 0 }
      'picture' { 2 }
      'comboBox' { 3 }
      'dropDownList' { 4 }
      'date' { 6 }
      'group' { 7 }
      'checkBox' { 8 }
      'repeatingSection' { 9 }
      default { 1 }
    }
  }
  function Test-ContentControlSelector($control, $selector) {
    if ($null -eq $selector) { return $true }
    if ($selector.id -and [long]$control.ID -ne [long]$selector.id) { return $false }
    if ($selector.tag -and [string]$control.Tag -ne [string]$selector.tag) { return $false }
    if ($selector.title -and [string]$control.Title -ne [string]$selector.title) { return $false }
    return $true
  }
  function Get-ContentControlSnapshot() {
    $controls = @()
    foreach ($control in $doc.ContentControls) {
      $entries = @()
      if ([int]$control.Type -in @(3, 4)) {
        foreach ($entry in $control.DropdownListEntries) { $entries += [pscustomobject]@{ index = [int]$entry.Index; text = [string]$entry.Text; value = [string]$entry.Value } }
      }
      $controls += [pscustomobject]@{
        id = [long]$control.ID
        type = [int]$control.Type
        typeName = Get-ContentControlTypeName ([int]$control.Type)
        tag = [string]$control.Tag
        title = [string]$control.Title
        text = $(try { [string]$control.Range.Text.Trim([char]13, [char]7) } catch { '' })
        checked = $(if ([int]$control.Type -eq 8) { try { [bool]$control.Checked } catch { $false } } else { $null })
        lockContents = $(try { [bool]$control.LockContents } catch { $false })
        lockControl = $(try { [bool]$control.LockContentControl } catch { $false })
        start = [int]$control.Range.Start
        end = [int]$control.Range.End
        entries = $entries
      }
    }
    return $controls
  }
  function Get-ControlValueProperty($control) {
    $keys = @()
    if ($control.Tag) { $keys += [string]$control.Tag }
    if ($control.Title) { $keys += [string]$control.Title }
    $keys += [string]$control.ID
    foreach ($key in $keys) {
      $mappedKey = if ($actionParams.fieldMap -and $actionParams.fieldMap.PSObject.Properties[$key]) { [string]$actionParams.fieldMap.PSObject.Properties[$key].Value } else { $key }
      $property = $actionParams.values.PSObject.Properties[$mappedKey]
      if ($null -ne $property) { return $property }
    }
    return $null
  }
  function Set-ContentControlValue($control, $rawValue) {
    $config = if ($rawValue -is [pscustomobject]) { $rawValue } else { [pscustomobject]@{ value = $rawValue } }
    $value = $config.value
    $type = [int]$control.Type
    try { $control.LockContents = $false } catch {}
    $fontName = $(try { [string]$control.Range.Font.Name } catch { '' })
    $fontSize = $(try { [double]$control.Range.Font.Size } catch { 0 })
    $fontBold = $(try { [int]$control.Range.Font.Bold } catch { 0 })
    $fontItalic = $(try { [int]$control.Range.Font.Italic } catch { 0 })
    $fontColor = $(try { [int]$control.Range.Font.Color } catch { 0 })
    if ($type -eq 8) {
      $control.Checked = if ($value -is [bool]) { [bool]$value } else { [string]$value -in @('1', 'true', 'True', 'yes', '是', '勾选') }
    } elseif ($type -in @(3, 4)) {
      $selected = $false
      foreach ($entry in $control.DropdownListEntries) {
        if ([string]$entry.Text -eq [string]$value -or [string]$entry.Value -eq [string]$value) { $entry.Select(); $selected = $true; break }
      }
      if (-not $selected -and $type -eq 3 -and $actionParams.allowNewListEntry -eq $true) {
        $entry = $control.DropdownListEntries.Add([string]$value, [string]$value)
        $entry.Select()
      } elseif (-not $selected -and $actionParams.strictListValues -eq $true) { throw '下拉控件不存在选项: ' + [string]$value }
    } elseif ($type -eq 2 -or $config.kind -eq 'image') {
      $imagePath = [string]$value
      if (-not [IO.File]::Exists($imagePath)) { throw '内容控件图片不存在: ' + $imagePath }
      $picture = $control.Range.InlineShapes.AddPicture($imagePath, $false, $true)
      if ($config.width) { $picture.LockAspectRatio = -1; $picture.Width = [double]$config.width }
      if ($config.height) { $picture.LockAspectRatio = -1; $picture.Height = [double]$config.height }
    } else {
      if ($type -eq 6 -and $config.dateFormat) { try { $control.DateDisplayFormat = [string]$config.dateFormat } catch {} }
      $control.Range.Text = [string]$value
      if ($actionParams.preserveFormatting -ne $false) {
        try { if ($fontName) { $control.Range.Font.Name = $fontName }; if ($fontSize -gt 0) { $control.Range.Font.Size = $fontSize }; $control.Range.Font.Bold = $fontBold; $control.Range.Font.Italic = $fontItalic; $control.Range.Font.Color = $fontColor } catch {}
      }
    }
    if ($null -ne $config.lockContents) { $control.LockContents = [bool]$config.lockContents }
    elseif ($actionParams.relock -eq $true) { $control.LockContents = $true }
  }
  function Add-ContentControl($config) {
    $type = Resolve-ContentControlType ([string]$config.type)
    if ($type -in @(8, 9) -and [int]$doc.CompatibilityMode -lt 14) {
      if ($actionParams.upgradeCompatibility -eq $false) {
        throw '复选框和重复节内容控件需要 Word 2010 或更高兼容模式'
      }
      $previousCompatibilityMode = [int]$doc.CompatibilityMode
      $doc.Convert()
      $operationData.compatibilityUpgrade = [pscustomobject]@{
        from = $previousCompatibilityMode
        to = [int]$doc.CompatibilityMode
      }
    }
    $start = if ($null -ne $config.start) { [int]$config.start } else { [Math]::Max(0, $doc.Content.End - 1) }
    $end = if ($null -ne $config.end) { [int]$config.end } else { $start }
    $range = $doc.Range([Math]::Max(0, $start), [Math]::Max($start, $end))
    $control = $doc.ContentControls.Add($type, $range)
    if ($config.tag) { $control.Tag = [string]$config.tag }
    if ($config.title) { $control.Title = [string]$config.title }
    if ($config.placeholder) { try { $control.SetPlaceholderText([Type]::Missing, [Type]::Missing, [string]$config.placeholder) } catch {} }
    if ($config.color) { try { $control.Color = [int]$config.color } catch {} }
    foreach ($entry in @($config.entries)) {
      if ($entry -is [string]) { [void]$control.DropdownListEntries.Add([string]$entry, [string]$entry) }
      elseif ($entry.text) { [void]$control.DropdownListEntries.Add([string]$entry.text, $(if ($null -ne $entry.value) { [string]$entry.value } else { [string]$entry.text })) }
    }
    if ($null -ne $config.value) { Set-ContentControlValue $control $config }
    if ($null -ne $config.lockContents) { $control.LockContents = [bool]$config.lockContents }
    if ($null -ne $config.lockControl) { $control.LockContentControl = [bool]$config.lockControl }
    return $control
  }

  if ($_operation -eq 'inspectContentControls') {
    $operationData.controls = @(Get-ContentControlSnapshot)
    $operationData.controlCount = @($operationData.controls).Count
    $changes = @()
  } elseif ($_operation -eq 'populateContentControls') {
    if ($null -eq $actionParams.values) { throw 'populateContentControls 需要 params.values 对象' }
    $filled = @(); $missing = @()
    foreach ($control in $doc.ContentControls) {
      $property = Get-ControlValueProperty $control
      if ($null -eq $property) {
        if ($actionParams.reportMissing -eq $true) { $missing += [pscustomobject]@{ id = [long]$control.ID; tag = [string]$control.Tag; title = [string]$control.Title } }
        continue
      }
      Set-ContentControlValue $control $property.Value
      $filled += [pscustomobject]@{ id = [long]$control.ID; tag = [string]$control.Tag; title = [string]$control.Title; type = Get-ContentControlTypeName ([int]$control.Type) }
    }
    $operationData.filled = $filled
    $operationData.filledCount = $filled.Count
    $operationData.missing = $missing
    $changes += [pscustomobject]@{ kind = 'content-control'; target = 'document'; detail = '已按控件类型填充模板字段' }
  } else {
    $command = if ($actionParams.command) { [string]$actionParams.command } else { 'add' }
    $affected = 0
    if ($command -eq 'add') {
      foreach ($config in @($actionParams.controls)) { $null = Add-ContentControl $config; $affected++ }
      if ($affected -eq 0) { $null = Add-ContentControl $actionParams; $affected = 1 }
    } else {
      for ($index = $doc.ContentControls.Count; $index -ge 1; $index--) {
        $control = $doc.ContentControls.Item($index)
        if (-not (Test-ContentControlSelector $control $actionParams.selector)) { continue }
        switch ($command) {
          'delete' { $control.Delete($actionParams.deleteContents -eq $true); $affected++ }
          'setLock' { if ($null -ne $actionParams.lockContents) { $control.LockContents = [bool]$actionParams.lockContents }; if ($null -ne $actionParams.lockControl) { $control.LockContentControl = [bool]$actionParams.lockControl }; $affected++ }
          'addListEntry' { [void]$control.DropdownListEntries.Add([string]$actionParams.text, $(if ($null -ne $actionParams.value) { [string]$actionParams.value } else { [string]$actionParams.text })); $affected++ }
          'clearListEntries' { $control.DropdownListEntries.Clear(); $affected++ }
          'setValue' { Set-ContentControlValue $control $actionParams; $affected++ }
          default { throw '不支持的内容控件命令: ' + $command }
        }
      }
    }
    $operationData.command = $command
    $operationData.affected = $affected
    $operationData.controls = @(Get-ContentControlSnapshot)
    $changes += [pscustomobject]@{ kind = 'content-control'; target = $command; detail = '已管理内容控件、列表项和锁定状态' }
  }
`;
}
