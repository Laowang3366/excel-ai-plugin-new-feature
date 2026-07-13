const POWER_QUERY_OPERATIONS = new Set([
  "createPowerQuery",
  "inspectPowerQueries",
  "managePowerQuery",
]);

export function isPowerQueryOperation(operation: string): boolean {
  return POWER_QUERY_OPERATIONS.has(operation);
}

export function buildPowerQueryOperationScript(operation: string): string | undefined {
  if (!isPowerQueryOperation(operation)) return undefined;
  return `
  function Get-WorkbookQuery([string]$name) {
    try { return $workbook.Queries.Item($name) } catch { return $null }
  }
  function Get-PowerQueryLoads([string]$name) {
    $loads = @()
    foreach ($ws in $workbook.Worksheets) {
      foreach ($table in $ws.ListObjects) {
        try {
          $queryTable = $table.QueryTable
          $connectionName = try { [string]$queryTable.WorkbookConnection.Name } catch { '' }
          $commandText = try { [string](@($queryTable.CommandText) -join '') } catch { '' }
          if ($connectionName -ieq ('Query - ' + $name) -or $commandText.Contains('[' + $name + ']')) {
            $loads += [pscustomobject]@{
              kind = 'worksheet'
              sheet = [string]$ws.Name
              table = [string]$table.Name
              range = [string]$table.Range.Address($false, $false)
              connection = $connectionName
              refreshing = $(try { [bool]$queryTable.Refreshing } catch { $false })
            }
          }
        } catch {}
      }
    }
    foreach ($connection in $workbook.Connections) {
      if ([string]$connection.Name -ieq ('Query - ' + $name)) {
        $inModel = try { $null -ne $connection.ModelConnection } catch { $false }
        $loads += [pscustomobject]@{
          kind = $(if ($inModel) { 'dataModel' } else { 'connection' })
          connection = [string]$connection.Name
        }
      }
    }
    return @($loads)
  }
  function Get-PowerQuerySnapshot([string]$onlyName) {
    $queries = @()
    try {
      foreach ($query in $workbook.Queries) {
        if ($onlyName -and [string]$query.Name -ine $onlyName) { continue }
        $queries += [pscustomobject]@{
          name = [string]$query.Name
          formula = [string]$query.Formula
          description = $(try { [string]$query.Description } catch { '' })
          loads = @(Get-PowerQueryLoads ([string]$query.Name))
        }
      }
    } catch { throw '当前 Excel/WPS 版本未提供 Workbook.Queries COM 接口' }
    $connections = @()
    foreach ($connection in $workbook.Connections) {
      if (-not $onlyName -or [string]$connection.Name -ieq ('Query - ' + $onlyName)) {
        $connections += [pscustomobject]@{
          name = [string]$connection.Name
          type = $(try { [int]$connection.Type } catch { $null })
          description = $(try { [string]$connection.Description } catch { '' })
          refreshWithRefreshAll = $(try { [bool]$connection.OLEDBConnection.RefreshWithRefreshAll } catch { $null })
        }
      }
    }
    return [pscustomobject]@{ queries = @($queries); connections = @($connections) }
  }
  function Remove-PowerQueryLoads([string]$name, [bool]$clearOutput) {
    $removed = 0
    foreach ($ws in $workbook.Worksheets) {
      foreach ($table in @($ws.ListObjects)) {
        try {
          $queryTable = $table.QueryTable
          $connectionName = try { [string]$queryTable.WorkbookConnection.Name } catch { '' }
          $commandText = try { [string](@($queryTable.CommandText) -join '') } catch { '' }
          if ($connectionName -ieq ('Query - ' + $name) -or $commandText.Contains('[' + $name + ']')) {
            $outputRange = $table.Range
            $table.Unlist()
            if ($clearOutput) { $outputRange.Clear() }
            $removed++
          }
        } catch {}
      }
    }
    foreach ($connection in @($workbook.Connections)) {
      if ([string]$connection.Name -ieq ('Query - ' + $name)) {
        try { $connection.Delete(); $removed++ } catch {}
      }
    }
    return $removed
  }
  function New-PowerQueryConnection([string]$name, [bool]$dataModel) {
    $connectionName = 'Query - ' + $name
    try { $workbook.Connections.Item($connectionName).Delete() } catch {}
    $connectionString = 'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=' + $name + ';Extended Properties=""'
    return $workbook.Connections.Add2(
      $connectionName,
      'Power Query - ' + $name,
      $connectionString,
      'SELECT * FROM [' + $name.Replace(']', ']]') + ']',
      2,
      $dataModel,
      $false
    )
  }
  function Add-PowerQueryWorksheetLoad([string]$name, [string]$destination, [string]$tableName) {
    if (-not $destination) { throw '加载到工作表需要 params.destination，例如 QueryOutput!A1' }
    $parts = $destination -split '!', 2
    if ($parts.Count -eq 2) {
      $targetSheet = $workbook.Worksheets.Item($parts[0].Trim("'"))
      $targetAddress = $parts[1]
    } else {
      $targetSheet = $sheet
      $targetAddress = $destination
    }
    $connectionString = 'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=' + $name + ';Extended Properties=""'
    $missing = [Type]::Missing
    $table = $targetSheet.ListObjects.Add(0, @($connectionString), $missing, 1, $targetSheet.Range($targetAddress))
    if ($tableName) { try { $table.Name = $tableName } catch {} }
    $queryTable = $table.QueryTable
    $queryTable.CommandType = 2
    $queryTable.CommandText = @('SELECT * FROM [' + $name.Replace(']', ']]') + ']')
    $queryTable.Refresh($false)
    return $table
  }

  if ($_operation -eq 'inspectPowerQueries') {
    $snapshot = Get-PowerQuerySnapshot ([string]$actionParams.name)
    $operationData.queries = $snapshot.queries
    $operationData.connections = $snapshot.connections
    $operationData.queryCount = @($snapshot.queries).Count
    $changes = @()
  } else {
    $command = if ($_operation -eq 'createPowerQuery') { 'upsert' } elseif ($actionParams.command) { [string]$actionParams.command } else { 'upsert' }
    $queryName = [string]$actionParams.name
    if (-not $queryName) { throw 'Power Query 操作需要 params.name' }
    $query = Get-WorkbookQuery $queryName
    switch ($command) {
      { $_ -in @('create', 'update', 'upsert') } {
        if (-not $actionParams.mFormula) { throw "$command 需要 params.mFormula" }
        if ($command -eq 'create' -and $null -ne $query) { throw "Power Query 已存在: $queryName" }
        if ($command -eq 'update' -and $null -eq $query) { throw "找不到 Power Query: $queryName" }
        if ($null -eq $query) {
          $query = $workbook.Queries.Add($queryName, [string]$actionParams.mFormula, [string]$actionParams.description)
        } else {
          $query.Formula = [string]$actionParams.mFormula
          if ($null -ne $actionParams.description) { try { $query.Description = [string]$actionParams.description } catch {} }
        }
      }
      'duplicate' {
        if ($null -eq $query) { throw "找不到 Power Query: $queryName" }
        $newName = [string]$actionParams.newName
        if (-not $newName) { throw 'duplicate 需要 params.newName' }
        if ($null -ne (Get-WorkbookQuery $newName)) { throw "Power Query 已存在: $newName" }
        $query = $workbook.Queries.Add($newName, [string]$query.Formula, [string]$query.Description)
        $queryName = $newName
      }
      'rename' {
        if ($null -eq $query) { throw "找不到 Power Query: $queryName" }
        $newName = [string]$actionParams.newName
        if (-not $newName) { throw 'rename 需要 params.newName' }
        if ($null -ne (Get-WorkbookQuery $newName)) { throw "Power Query 已存在: $newName" }
        $previousLoads = @(Get-PowerQueryLoads $queryName)
        $worksheetLoads = @($previousLoads | Where-Object { $_.kind -eq 'worksheet' })
        $hadModelLoad = @($previousLoads | Where-Object { $_.kind -eq 'dataModel' }).Count -gt 0
        $hadConnectionLoad = @($previousLoads | Where-Object { $_.kind -eq 'connection' }).Count -gt 0
        if ($previousLoads.Count -gt 0) { [void](Remove-PowerQueryLoads $queryName $true) }
        $query.Name = $newName
        $queryName = $newName
        foreach ($previousLoad in $worksheetLoads) {
          $loadSheet = $workbook.Worksheets.Item([string]$previousLoad.sheet)
          $loadStart = $loadSheet.Range([string]$previousLoad.range).Cells.Item(1, 1).Address($false, $false)
          [void](Add-PowerQueryWorksheetLoad $queryName ([string]$previousLoad.sheet + '!' + $loadStart) ([string]$previousLoad.table))
        }
        if ($hadModelLoad) { [void](New-PowerQueryConnection $queryName $true) }
        elseif ($hadConnectionLoad -and $worksheetLoads.Count -eq 0) { [void](New-PowerQueryConnection $queryName $false) }
      }
      'load' {
        if ($null -eq $query) { throw "找不到 Power Query: $queryName" }
        if ($actionParams.replaceLoad -ne $false) { [void](Remove-PowerQueryLoads $queryName ([bool]$actionParams.clearOutput)) }
        $loadMode = if ($actionParams.loadMode) { [string]$actionParams.loadMode } else { 'worksheet' }
        if ($loadMode -eq 'dataModel') {
          [void](New-PowerQueryConnection $queryName $true)
        } elseif ($loadMode -eq 'connectionOnly') {
          [void](New-PowerQueryConnection $queryName $false)
        } else {
          [void](Add-PowerQueryWorksheetLoad $queryName ([string]$actionParams.destination) ([string]$actionParams.tableName))
        }
      }
      'refresh' {
        if ($null -eq $query) { throw "找不到 Power Query: $queryName" }
        $refreshed = 0
        foreach ($connection in $workbook.Connections) {
          if ([string]$connection.Name -ieq ('Query - ' + $queryName)) { $connection.Refresh(); $refreshed++ }
        }
        foreach ($ws in $workbook.Worksheets) {
          foreach ($table in $ws.ListObjects) {
            try {
              if ([string]$table.QueryTable.WorkbookConnection.Name -ieq ('Query - ' + $queryName)) {
                $table.QueryTable.Refresh($false); $refreshed++
              }
            } catch {}
          }
        }
        try { $app.CalculateUntilAsyncQueriesDone() } catch {}
        $operationData.refreshedLoads = $refreshed
      }
      'unload' {
        if ($null -eq $query) { throw "找不到 Power Query: $queryName" }
        $operationData.removedLoads = Remove-PowerQueryLoads $queryName ($actionParams.clearOutput -eq $true)
      }
      'delete' {
        if ($null -eq $query) { throw "找不到 Power Query: $queryName" }
        $operationData.removedLoads = Remove-PowerQueryLoads $queryName ($actionParams.clearOutput -eq $true)
        $query.Delete()
      }
      default { throw "不支持的 Power Query 命令: $command" }
    }

    if ($command -notin @('load', 'refresh', 'unload', 'delete') -and $actionParams.loadMode) {
      if ($actionParams.replaceLoad -ne $false) { [void](Remove-PowerQueryLoads $queryName ([bool]$actionParams.clearOutput)) }
      if ([string]$actionParams.loadMode -eq 'dataModel') {
        [void](New-PowerQueryConnection $queryName $true)
      } elseif ([string]$actionParams.loadMode -eq 'connectionOnly') {
        [void](New-PowerQueryConnection $queryName $false)
      } elseif ([string]$actionParams.loadMode -eq 'worksheet') {
        [void](Add-PowerQueryWorksheetLoad $queryName ([string]$actionParams.destination) ([string]$actionParams.tableName))
      }
    }
    if ($actionParams.refresh -eq $true) {
      foreach ($connection in $workbook.Connections) { if ([string]$connection.Name -ieq ('Query - ' + $queryName)) { $connection.Refresh() } }
      try { $app.CalculateUntilAsyncQueriesDone() } catch {}
    }
    $snapshot = Get-PowerQuerySnapshot $(if ($command -eq 'delete') { '' } else { $queryName })
    $operationData.command = $command
    $operationData.queryName = $queryName
    $operationData.queries = $snapshot.queries
    $operationData.connections = $snapshot.connections
    $changes += [pscustomobject]@{ kind = 'power-query'; target = $queryName; detail = '已执行 Power Query ' + $command }
  }
`;
}
