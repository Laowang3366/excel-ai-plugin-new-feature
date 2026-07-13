const FORMULA_OPERATIONS = new Set([
  "traceFormulaDependencies",
  "inspectFormulaDependencies",
  "repairFormulaReferences",
  "convertFormulasToValues",
  "inspectFormulaBackups",
  "restoreFormulas",
  "inspectFormulaProtection",
  "manageFormulaProtection",
]);

export function isExcelFormulaGovernanceOperation(operation: string): boolean {
  return FORMULA_OPERATIONS.has(operation);
}

export function buildExcelFormulaGovernanceScript(operation: string): string | undefined {
  if (!isExcelFormulaGovernanceOperation(operation)) return undefined;
  return String.raw`
  function Get-FormulaText($cell) {
    try { return [string]$cell.Formula2 } catch { return [string]$cell.Formula }
  }
  function Get-FormulaR1C1Text($cell) {
    try { return [string]$cell.Formula2R1C1 } catch { return [string]$cell.FormulaR1C1 }
  }
  function Get-CellId($cell) {
    return [string]$cell.Worksheet.Name + '!' + [string]$cell.Address($false, $false)
  }
  function Get-FormulaScopes([string]$defaultScope) {
    $scope = if ($actionParams.scope) { [string]$actionParams.scope } else { $defaultScope }
    $scopes = @()
    if ($scope -eq 'workbook') {
      foreach ($ws in $workbook.Worksheets) { $scopes += [pscustomobject]@{ worksheet = $ws; target = $ws.UsedRange } }
    } elseif ($scope -eq 'sheet') {
      $scopes += [pscustomobject]@{ worksheet = $sheet; target = $sheet.UsedRange }
    } else {
      $scopes += [pscustomobject]@{ worksheet = $sheet; target = $range }
    }
    return @($scopes)
  }
  function Get-FormulaCells($targetRange) {
    try { return @($targetRange.SpecialCells(-4123).Cells) } catch {}
    $fallback = @()
    foreach ($candidate in $targetRange.Cells) {
      try { if ($candidate.HasFormula) { $fallback += $candidate } } catch {}
    }
    return @($fallback)
  }
  function Add-DependencyEdge($from, $to, $kind, $reference, $edges, $edgeKeys) {
    if (-not $to) { return }
    $key = $from + '|' + $to + '|' + $kind
    if ($edgeKeys.ContainsKey($key)) { return }
    $edgeKeys[$key] = $true
    [void]$edges.Add([pscustomobject]@{ from = $from; to = $to; kind = $kind; reference = $reference })
  }
  function Add-RangeDependencies($from, $referenceRange, $kind, $reference, $edges, $edgeKeys, [int]$maxCells) {
    if ($null -eq $referenceRange) { return }
    $count = $(try { [int]$referenceRange.Cells.Count } catch { 0 })
    if ($count -gt $maxCells) {
      $rangeId = [string]$referenceRange.Worksheet.Name + '!' + [string]$referenceRange.Address($false, $false)
      Add-DependencyEdge $from $rangeId ($kind + '-range') $reference $edges $edgeKeys
      return
    }
    foreach ($referencedCell in $referenceRange.Cells) {
      Add-DependencyEdge $from (Get-CellId $referencedCell) $kind $reference $edges $edgeKeys
    }
  }
  function Resolve-FormulaReferences($cell, [string]$formula, $edges, $edgeKeys, $broken, [int]$maxCells) {
    $from = Get-CellId $cell
    if ($formula.Contains('#REF!')) { [void]$broken.Add([pscustomobject]@{ cell = $from; formula = $formula; reason = '#REF!' }) }
    $withoutStrings = [regex]::Replace($formula, '"(?:[^"]|"")*"', '""')
    $externalPattern = '\[[^\]]+\](?:''(?<sheet>(?:[^'']|'''')+)''|(?<sheet>[^!]+))!(?<address>\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)'
    foreach ($match in [regex]::Matches($withoutStrings, $externalPattern)) {
      Add-DependencyEdge $from ('external:' + $match.Value) 'external' $match.Value $edges $edgeKeys
    }
    $withoutExternal = [regex]::Replace($withoutStrings, $externalPattern, '')
    $qualifiedPattern = '(?<![A-Za-z0-9_.])(?:''(?<quoted>(?:[^'']|'''')+)''|(?<plain>[A-Za-z_\u4e00-\u9fff][A-Za-z0-9_ .\u4e00-\u9fff-]*))!(?<address>\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)(?<spill>#?)'
    foreach ($match in [regex]::Matches($withoutExternal, $qualifiedPattern)) {
      $sheetName = if ($match.Groups['quoted'].Success) { $match.Groups['quoted'].Value.Replace("''", "'") } else { $match.Groups['plain'].Value.Trim() }
      $address = $match.Groups['address'].Value
      try {
        $refSheet = $workbook.Worksheets.Item($sheetName)
        $refRange = $refSheet.Range($address)
        Add-RangeDependencies $from $refRange 'cross-sheet' $match.Value $edges $edgeKeys $maxCells
      } catch {
        [void]$broken.Add([pscustomobject]@{ cell = $from; formula = $formula; reason = '找不到跨表引用: ' + $match.Value })
      }
    }
    $localFormula = [regex]::Replace($withoutExternal, $qualifiedPattern, '')
    $localPattern = '(?<![A-Za-z0-9_.!])(?<address>\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)(?<spill>#?)'
    foreach ($match in [regex]::Matches($localFormula, $localPattern)) {
      try {
        $refRange = $cell.Worksheet.Range($match.Groups['address'].Value)
        Add-RangeDependencies $from $refRange 'same-sheet' $match.Value $edges $edgeKeys $maxCells
      } catch {}
    }
    foreach ($definedName in $workbook.Names) {
      $shortName = ([string]$definedName.Name -split '!', 2)[-1].Trim("'")
      if (-not $shortName) { continue }
      if ([regex]::IsMatch($withoutStrings, '(?<![A-Za-z0-9_.])' + [regex]::Escape($shortName) + '(?![A-Za-z0-9_.])', 1)) {
        try { Add-RangeDependencies $from $definedName.RefersToRange 'defined-name' $shortName $edges $edgeKeys $maxCells } catch {}
      }
    }
    $structuredPattern = '(?<![A-Za-z0-9_.])(?<table>[A-Za-z_][A-Za-z0-9_.]*)\[(?<column>[^\]]+)\]'
    foreach ($match in [regex]::Matches($withoutStrings, $structuredPattern)) {
      $tableName = $match.Groups['table'].Value
      $columnName = $match.Groups['column'].Value.Trim('[', ']', '@', '#')
      foreach ($ws in $workbook.Worksheets) {
        try {
          $table = $ws.ListObjects.Item($tableName)
          $refRange = if ($columnName -and $columnName -notin @('All', 'Data', 'Headers', 'Totals', 'This Row')) { $table.ListColumns.Item($columnName).Range } else { $table.Range }
          Add-RangeDependencies $from $refRange 'structured-reference' $match.Value $edges $edgeKeys $maxCells
          break
        } catch {}
      }
    }
  }
  function Visit-FormulaNode([string]$nodeId, $graph, $states, $stack, $cycles, $cycleKeys) {
    $states[$nodeId] = 1
    [void]$stack.Add($nodeId)
    foreach ($dependency in @($graph[$nodeId])) {
      if (-not $graph.ContainsKey($dependency)) { continue }
      $state = if ($states.ContainsKey($dependency)) { [int]$states[$dependency] } else { 0 }
      if ($state -eq 0) {
        Visit-FormulaNode $dependency $graph $states $stack $cycles $cycleKeys
      } elseif ($state -eq 1) {
        $start = $stack.IndexOf($dependency)
        if ($start -ge 0) {
          $cycle = @($stack.GetRange($start, $stack.Count - $start).ToArray()) + @($dependency)
          $key = $cycle -join ' -> '
          if (-not $cycleKeys.ContainsKey($key)) { $cycleKeys[$key] = $true; [void]$cycles.Add($cycle) }
        }
      }
    }
    [void]$stack.RemoveAt($stack.Count - 1)
    $states[$nodeId] = 2
  }
  function Build-FormulaDependencyReport() {
    $defaultScope = if ($targetExplicit) { 'target' } else { 'workbook' }
    $scopes = @(Get-FormulaScopes $defaultScope)
    $nodeMap = @{}
    foreach ($scopeItem in $scopes) {
      foreach ($formulaCell in @(Get-FormulaCells $scopeItem.target)) {
        $id = Get-CellId $formulaCell
        $formula = Get-FormulaText $formulaCell
        $nodeMap[$id] = [pscustomobject]@{
          id = $id
          sheet = [string]$formulaCell.Worksheet.Name
          address = [string]$formulaCell.Address($false, $false)
          formula = $formula
          value = $(try { $formulaCell.Value2 } catch { $null })
          displayValue = $(try { [string]$formulaCell.Text } catch { '' })
          precedents = New-Object System.Collections.Generic.List[string]
          dependents = New-Object System.Collections.Generic.List[string]
        }
      }
    }
    $edges = New-Object System.Collections.Generic.List[object]
    $edgeKeys = @{}
    $broken = New-Object System.Collections.Generic.List[object]
    $maxCells = if ($actionParams.maxExpandedRangeCells) { [Math]::Max(1, [int]$actionParams.maxExpandedRangeCells) } else { 500 }
    foreach ($node in $nodeMap.Values) {
      $ws = $workbook.Worksheets.Item([string]$node.sheet)
      $cell = $ws.Range([string]$node.address)
      Resolve-FormulaReferences $cell ([string]$node.formula) $edges $edgeKeys $broken $maxCells
    }
    $graph = @{}
    foreach ($id in $nodeMap.Keys) { $graph[$id] = New-Object System.Collections.Generic.List[string] }
    foreach ($edge in $edges) {
      if ($nodeMap.ContainsKey([string]$edge.from)) { [void]$nodeMap[[string]$edge.from].precedents.Add([string]$edge.to) }
      if ($nodeMap.ContainsKey([string]$edge.to)) { [void]$nodeMap[[string]$edge.to].dependents.Add([string]$edge.from) }
      if ($graph.ContainsKey([string]$edge.from)) { [void]$graph[[string]$edge.from].Add([string]$edge.to) }
    }
    $cycles = New-Object System.Collections.Generic.List[object]
    $cycleKeys = @{}
    $states = @{}
    $stack = New-Object System.Collections.Generic.List[string]
    foreach ($id in $nodeMap.Keys) { if (-not $states.ContainsKey($id)) { Visit-FormulaNode $id $graph $states $stack $cycles $cycleKeys } }
    $excelCircular = $null
    try { if ($null -ne $app.CircularReference) { $excelCircular = Get-CellId $app.CircularReference } } catch {}
    $outputNodes = @()
    foreach ($node in $nodeMap.Values) {
      $outputNodes += [pscustomobject]@{
        id = $node.id
        sheet = $node.sheet
        address = $node.address
        formula = $node.formula
        value = $node.value
        displayValue = $node.displayValue
        precedents = @($node.precedents | ForEach-Object { $_ })
        dependents = @($node.dependents | ForEach-Object { $_ })
      }
    }
    $outputCycles = @()
    foreach ($cycle in $cycles) {
      $outputCycles += [pscustomobject]@{ path = @($cycle | ForEach-Object { $_ }) }
    }
    return [pscustomobject]@{
      nodes = $outputNodes
      edges = @($edges | ForEach-Object { $_ })
      cycles = $outputCycles
      circularReference = $excelCircular
      brokenReferences = @($broken | ForEach-Object { $_ })
      formulaCount = $nodeMap.Count
      edgeCount = $edges.Count
    }
  }
  function Get-FormulaBackupSheet([bool]$create) {
    $prefix = '_WenggeFormulaBackup'
    foreach ($ws in $workbook.Worksheets) {
      if ([string]$ws.Name -like ($prefix + '*') -and [string]$ws.Range('A1').Value2 -eq 'WENGGE_FORMULA_BACKUP_V1') {
        Write-Output -NoEnumerate $ws
        return
      }
    }
    if (-not $create) { return $null }
    $missing = [Type]::Missing
    $backupSheet = $workbook.Worksheets.Add($missing, $workbook.Worksheets.Item($workbook.Worksheets.Count), 1, -4167)
    $name = $prefix; $suffix = 1
    while ($true) { try { $backupSheet.Name = $name; break } catch { $suffix++; $name = $prefix + $suffix } }
    $backupSheet.Range('A1').Value2 = 'WENGGE_FORMULA_BACKUP_V1'
    $headers = @('backupId', 'createdAt', 'sheet', 'address', 'formula', 'formulaR1C1', 'numberFormat', 'locked', 'spillAddress', 'sourceRange')
    for ($column = 1; $column -le $headers.Count; $column++) { $backupSheet.Cells.Item(2, $column).Value2 = $headers[$column - 1] }
    $backupSheet.Visible = 2
    Write-Output -NoEnumerate $backupSheet
  }
  function Get-BackupLastRow($backupSheet) {
    return [Math]::Max(2, [int]$backupSheet.Cells.Item($backupSheet.Rows.Count, 1).End(-4162).Row)
  }
  function Get-SpillAddress($cell) {
    try { return [string]$cell.SpillingToRange.Address($false, $false) } catch {}
    try {
      $arrayRange = $cell.CurrentArray
      if ($arrayRange.Cells.Count -gt 1) { return [string]$arrayRange.Address($false, $false) }
    } catch {}
    return ''
  }
  function Write-FormulaBackupRecord($backupSheet, [int]$row, [string]$backupId, [string]$createdAt, $cell, [string]$sourceRange) {
    $formula = Get-FormulaText $cell
    $formulaR1C1 = Get-FormulaR1C1Text $cell
    $values = @(
      $backupId,
      $createdAt,
      [string]$cell.Worksheet.Name,
      [string]$cell.Address($false, $false),
      $formula,
      $formulaR1C1,
      [string]$cell.NumberFormat,
      $(if ([bool]$cell.Locked) { '1' } else { '0' }),
      (Get-SpillAddress $cell),
      $sourceRange
    )
    for ($column = 1; $column -le $values.Count; $column++) {
      $target = $backupSheet.Cells.Item($row, $column)
      if ($column -in @(5, 6)) { $target.NumberFormat = '@' }
      $target.Value2 = $values[$column - 1]
    }
  }
  function Get-FormulaBackupSummary($backupSheet) {
    if ($null -eq $backupSheet) { return @() }
    $groups = @{}
    $lastRow = Get-BackupLastRow $backupSheet
    for ($row = 3; $row -le $lastRow; $row++) {
      $id = [string]$backupSheet.Cells.Item($row, 1).Value2
      if (-not $id) { continue }
      if (-not $groups.ContainsKey($id)) {
        $groups[$id] = [ordered]@{ backupId = $id; createdAt = [string]$backupSheet.Cells.Item($row, 2).Value2; formulaCount = 0; sheets = New-Object System.Collections.Generic.HashSet[string]; ranges = New-Object System.Collections.Generic.HashSet[string] }
      }
      $groups[$id].formulaCount++
      [void]$groups[$id].sheets.Add([string]$backupSheet.Cells.Item($row, 3).Value2)
      [void]$groups[$id].ranges.Add([string]$backupSheet.Cells.Item($row, 10).Value2)
    }
    $summaries = @()
    foreach ($group in $groups.Values) {
      $summaries += [pscustomobject]@{
        backupId = $group.backupId
        createdAt = $group.createdAt
        formulaCount = $group.formulaCount
        sheets = @($group.sheets | ForEach-Object { $_ })
        ranges = @($group.ranges | Where-Object { $_ } | ForEach-Object { $_ })
      }
    }
    return @($summaries | Sort-Object createdAt -Descending)
  }
  function Get-FormulaProtectionSnapshot($scopes) {
    $items = @()
    foreach ($scopeItem in $scopes) {
      $formulaCells = @(Get-FormulaCells $scopeItem.target)
      $locked = 0
      foreach ($formulaCell in $formulaCells) { if ([bool]$formulaCell.Locked) { $locked++ } }
      $items += [pscustomobject]@{
        sheet = [string]$scopeItem.worksheet.Name
        protected = [bool]$scopeItem.worksheet.ProtectContents
        target = [string]$scopeItem.target.Address($false, $false)
        formulaCount = $formulaCells.Count
        lockedFormulaCount = $locked
      }
    }
    return @($items)
  }

  if ($_operation -in @('traceFormulaDependencies', 'inspectFormulaDependencies')) {
    $report = Build-FormulaDependencyReport
    $operationData.nodes = $report.nodes
    $operationData.edges = $report.edges
    $operationData.cycles = $report.cycles
    $operationData.circularReference = $report.circularReference
    $operationData.brokenReferences = $report.brokenReferences
    $operationData.formulaCount = $report.formulaCount
    $operationData.edgeCount = $report.edgeCount
    $changes = @()
  } elseif ($_operation -eq 'repairFormulaReferences') {
    $scopes = @(Get-FormulaScopes $(if ($targetExplicit) { 'target' } else { 'workbook' }))
    $repairs = @(); $unresolved = @()
    foreach ($scopeItem in $scopes) {
      foreach ($cell in @(Get-FormulaCells $scopeItem.target)) {
        $before = Get-FormulaText $cell
        if (-not $before.Contains('#REF!') -and $actionParams.applyAllMappings -ne $true) { continue }
        $after = $before
        foreach ($replacement in @($actionParams.replacements)) {
          if ($null -eq $replacement -or -not $replacement.find) { continue }
          if ($replacement.sheetName -and [string]$replacement.sheetName -ine [string]$cell.Worksheet.Name) { continue }
          $after = $after.Replace([string]$replacement.find, [string]$replacement.replace)
        }
        $strategy = 'mapping'
        if ($after.Contains('#REF!') -and $actionParams.copyFromNeighbors -eq $true) {
          $neighbors = @()
          foreach ($offset in @(@(0, -1), @(0, 1), @(-1, 0), @(1, 0))) {
            try { $neighbor = $cell.Offset($offset[0], $offset[1]); if ($neighbor.HasFormula -and -not (Get-FormulaText $neighbor).Contains('#REF!')) { $neighbors += $neighbor } } catch {}
          }
          if ($neighbors.Count -gt 0) {
            try { $cell.FormulaR1C1 = Get-FormulaR1C1Text $neighbors[0]; $after = Get-FormulaText $cell; $strategy = 'neighbor-r1c1' } catch {}
          }
        }
        if ($after -ne $before -and $strategy -eq 'mapping') { try { $cell.Formula2 = $after } catch { $cell.Formula = $after } }
        $current = Get-FormulaText $cell
        $item = [pscustomobject]@{ cell = Get-CellId $cell; before = $before; after = $current; strategy = $strategy }
        if ($current.Contains('#REF!')) { $unresolved += $item } else { $repairs += $item }
      }
    }
    try { $app.Calculate() } catch {}
    $operationData.repairs = @($repairs)
    $operationData.repairedCount = @($repairs).Count
    $operationData.unresolved = @($unresolved)
    $operationData.unresolvedCount = @($unresolved).Count
    $changes += [pscustomobject]@{ kind = 'formula-repair'; target = $_rangeAddress; detail = '已修复 ' + @($repairs).Count + ' 个错误引用' }
  } elseif ($_operation -eq 'convertFormulasToValues') {
    $scopes = @(Get-FormulaScopes $(if ($targetExplicit) { 'target' } else { 'sheet' }))
    $backupEnabled = $actionParams.createBackup -ne $false
    $backupId = if ($actionParams.backupId) { [string]$actionParams.backupId } else { [Guid]::NewGuid().ToString('N') }
    $createdAt = [DateTime]::UtcNow.ToString('o')
    $backupSheet = if ($backupEnabled) { Get-FormulaBackupSheet $true } else { $null }
    $nextRow = if ($null -ne $backupSheet) { (Get-BackupLastRow $backupSheet) + 1 } else { 0 }
    $formulaCount = 0; $convertedRanges = @()
    $convertedArrays = @{}
    foreach ($scopeItem in $scopes) {
      $formulaCells = @(Get-FormulaCells $scopeItem.target)
      foreach ($cell in $formulaCells) {
        if ($backupEnabled) { Write-FormulaBackupRecord $backupSheet $nextRow $backupId $createdAt $cell ([string]$scopeItem.target.Address($false, $false)); $nextRow++ }
        $formulaCount++
        $spillAddress = Get-SpillAddress $cell
        $spillKey = [string]$cell.Worksheet.Name + '!' + $spillAddress
        if ($spillAddress -and -not $convertedArrays.ContainsKey($spillKey)) {
          $spillRange = $cell.Worksheet.Range($spillAddress)
          $spillRange.Copy()
          $spillRange.PasteSpecial(-4163)
          $convertedArrays[$spillKey] = $true
        } elseif (-not $spillAddress) {
          $cell.Copy()
          $cell.PasteSpecial(-4163)
        }
      }
      if ($formulaCells.Count -gt 0) {
        $convertedRanges += [string]$scopeItem.worksheet.Name + '!' + [string]$scopeItem.target.Address($false, $false)
      }
    }
    try { $app.CutCopyMode = 0 } catch {}
    if ($null -ne $backupSheet) { $backupSheet.Visible = 2 }
    $operationData.backupId = $(if ($backupEnabled) { $backupId } else { $null })
    $operationData.convertedFormulaCells = $formulaCount
    $operationData.convertedRanges = @($convertedRanges)
    $changes += [pscustomobject]@{ kind = 'formula-values'; target = $_rangeAddress; detail = '已备份并将 ' + $formulaCount + ' 个公式转换为值' }
  } elseif ($_operation -eq 'inspectFormulaBackups') {
    $backupSheet = Get-FormulaBackupSheet $false
    $operationData.backups = @(Get-FormulaBackupSummary $backupSheet)
    $operationData.backupCount = @($operationData.backups).Count
    $changes = @()
  } elseif ($_operation -eq 'restoreFormulas') {
    $backupSheet = Get-FormulaBackupSheet $false
    if ($null -eq $backupSheet) { throw '当前工作簿没有公式备份' }
    $summaries = @(Get-FormulaBackupSummary $backupSheet)
    $backupId = if ($actionParams.backupId) { [string]$actionParams.backupId } elseif ($summaries.Count -gt 0) { [string]$summaries[0].backupId } else { '' }
    if (-not $backupId) { throw '找不到可恢复的公式备份' }
    $lastRow = Get-BackupLastRow $backupSheet
    $rows = @(); $restored = @(); $failed = @()
    for ($row = 3; $row -le $lastRow; $row++) { if ([string]$backupSheet.Cells.Item($row, 1).Value2 -eq $backupId) { $rows += $row } }
    foreach ($row in $rows) {
      $sheetName = [string]$backupSheet.Cells.Item($row, 3).Value2
      $address = [string]$backupSheet.Cells.Item($row, 4).Value2
      if ($actionParams.sheetName -and [string]$actionParams.sheetName -ine $sheetName) { continue }
      try {
        $targetSheet = $workbook.Worksheets.Item($sheetName)
        $targetCell = $targetSheet.Range($address)
        $spillAddress = [string]$backupSheet.Cells.Item($row, 9).Value2
        if ($spillAddress) { try { $targetSheet.Range($spillAddress).ClearContents() } catch {} }
        $formula = [string]$backupSheet.Cells.Item($row, 5).Value2
        try { $targetCell.Formula2 = $formula } catch { $targetCell.Formula = $formula }
        $targetCell.NumberFormat = [string]$backupSheet.Cells.Item($row, 7).Value2
        $targetCell.Locked = [string]$backupSheet.Cells.Item($row, 8).Value2 -eq '1'
        $restored += [pscustomobject]@{ cell = $sheetName + '!' + $address; formula = $formula }
      } catch {
        $failed += [pscustomobject]@{ cell = $sheetName + '!' + $address; error = $_.Exception.Message }
      }
    }
    if ($actionParams.removeAfterRestore -eq $true) { foreach ($row in @($rows | Sort-Object -Descending)) { $backupSheet.Rows.Item($row).Delete() } }
    $backupSheet.Visible = 2
    try { $app.Calculate() } catch {}
    $operationData.backupId = $backupId
    $operationData.restored = @($restored)
    $operationData.restoredCount = @($restored).Count
    $operationData.failed = @($failed)
    $changes += [pscustomobject]@{ kind = 'formula-restore'; target = $backupId; detail = '已恢复 ' + @($restored).Count + ' 个公式' }
  } elseif ($_operation -eq 'inspectFormulaProtection') {
    $scopes = @(Get-FormulaScopes $(if ($targetExplicit) { 'target' } else { 'workbook' }))
    $operationData.protection = @(Get-FormulaProtectionSnapshot $scopes)
    $changes = @()
  } else {
    $command = if ($actionParams.command) { [string]$actionParams.command } else { 'lock' }
    $scopes = @(Get-FormulaScopes $(if ($targetExplicit) { 'target' } else { 'sheet' }))
    foreach ($scopeItem in $scopes) {
      $ws = $scopeItem.worksheet
      $password = [string]$actionParams.password
      if ($command -eq 'lock') {
        try { $ws.Unprotect($password) } catch {}
        if ($actionParams.unlockInputs -ne $false) { $scopeItem.target.Locked = $false }
        foreach ($formulaCell in @(Get-FormulaCells $scopeItem.target)) { $formulaCell.Locked = $true }
        if ($actionParams.protectSheet -ne $false) { $ws.Protect($password, $true, $true, $true, $true) }
      } elseif ($command -eq 'unlock') {
        $ws.Unprotect($password)
        foreach ($formulaCell in @(Get-FormulaCells $scopeItem.target)) { $formulaCell.Locked = $false }
      } else { throw '不支持的公式保护命令: ' + $command }
    }
    $operationData.command = $command
    $operationData.protection = @(Get-FormulaProtectionSnapshot $scopes)
    $changes += [pscustomobject]@{ kind = 'formula-protection'; target = $_rangeAddress; detail = '已执行公式区域保护命令: ' + $command }
  }
`;
}
