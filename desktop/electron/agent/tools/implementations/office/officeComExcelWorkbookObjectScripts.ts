const WORKBOOK_OBJECT_OPERATIONS = new Set([
  "inspectWorkbookObjects",
  "manageWorkbookObject",
  "manageWorksheetObjects",
]);

export function isWorkbookObjectOperation(operation: string): boolean {
  return WORKBOOK_OBJECT_OPERATIONS.has(operation);
}

export function buildWorkbookObjectOperationScript(operation: string): string | undefined {
  if (!isWorkbookObjectOperation(operation)) return undefined;
  return `
  function Convert-ObjectHexColor([object]$value, [int]$fallback) {
    if ($null -eq $value) { return $fallback }
    $hex = ([string]$value).Trim().TrimStart('#')
    if ($hex -notmatch '^[0-9a-fA-F]{6}$') { return $fallback }
    return [Convert]::ToInt32($hex.Substring(0, 2), 16) +
      (256 * [Convert]::ToInt32($hex.Substring(2, 2), 16)) +
      (65536 * [Convert]::ToInt32($hex.Substring(4, 2), 16))
  }
  function Include-ObjectType([string]$type) {
    $requested = @($actionParams.types | Where-Object { $_ })
    return $requested.Count -eq 0 -or $requested -contains $type
  }
  function Get-ObjectSheet() {
    $name = if ($actionParams.sheetName) { [string]$actionParams.sheetName } else { $_sheetName }
    try { return $workbook.Worksheets.Item($name) } catch { throw "找不到工作表: $name" }
  }
  function Get-WorkbookObjectSnapshot() {
    $result = [ordered]@{}
    if (Include-ObjectType 'worksheet') {
      $items = @()
      foreach ($ws in $workbook.Worksheets) {
        $items += [pscustomobject]@{
          name = [string]$ws.Name
          index = [int]$ws.Index
          visible = [int]$ws.Visible
          usedRange = $(try { [string]$ws.UsedRange.Address($false, $false) } catch { '' })
          tableCount = $(try { [int]$ws.ListObjects.Count } catch { 0 })
          chartCount = $(try { [int]$ws.ChartObjects().Count } catch { 0 })
          shapeCount = $(try { [int]$ws.Shapes.Count } catch { 0 })
        }
      }
      $result.worksheets = @($items)
    }
    if (Include-ObjectType 'name') {
      $items = @()
      foreach ($definedName in $workbook.Names) {
        $items += [pscustomobject]@{
          name = [string]$definedName.Name
          refersTo = $(try { [string]$definedName.RefersTo } catch { '' })
          visible = $(try { [bool]$definedName.Visible } catch { $true })
          scope = 'workbook'
        }
      }
      foreach ($ws in $workbook.Worksheets) {
        foreach ($definedName in $ws.Names) {
          $items += [pscustomobject]@{
            name = [string]$definedName.Name
            refersTo = $(try { [string]$definedName.RefersTo } catch { '' })
            visible = $(try { [bool]$definedName.Visible } catch { $true })
            scope = [string]$ws.Name
          }
        }
      }
      $result.names = @($items)
    }
    if (Include-ObjectType 'table') {
      $items = @()
      foreach ($ws in $workbook.Worksheets) {
        foreach ($table in $ws.ListObjects) {
          $items += [pscustomobject]@{
            sheet = [string]$ws.Name
            name = [string]$table.Name
            displayName = [string]$table.DisplayName
            range = [string]$table.Range.Address($false, $false)
            style = $(try { [string]$table.TableStyle } catch { '' })
            showTotals = [bool]$table.ShowTotals
            sourceType = [int]$table.SourceType
          }
        }
      }
      $result.tables = @($items)
    }
    if (Include-ObjectType 'chart') {
      $items = @()
      foreach ($ws in $workbook.Worksheets) {
        foreach ($chartObject in $ws.ChartObjects()) {
          $items += [pscustomobject]@{
            sheet = [string]$ws.Name
            name = [string]$chartObject.Name
            chartType = $(try { [int]$chartObject.Chart.ChartType } catch { $null })
            title = $(try { $(if ($chartObject.Chart.HasTitle) { [string]$chartObject.Chart.ChartTitle.Text } else { '' }) } catch { '' })
            left = [double]$chartObject.Left
            top = [double]$chartObject.Top
            width = [double]$chartObject.Width
            height = [double]$chartObject.Height
          }
        }
      }
      $result.charts = @($items)
    }
    if (Include-ObjectType 'shape') {
      $items = @()
      foreach ($ws in $workbook.Worksheets) {
        foreach ($shape in $ws.Shapes) {
          $items += [pscustomobject]@{
            sheet = [string]$ws.Name
            name = [string]$shape.Name
            type = [int]$shape.Type
            left = [double]$shape.Left
            top = [double]$shape.Top
            width = [double]$shape.Width
            height = [double]$shape.Height
            visible = [int]$shape.Visible
            alternativeText = $(try { [string]$shape.AlternativeText } catch { '' })
          }
        }
      }
      $result.shapes = @($items)
    }
    if (Include-ObjectType 'connection') {
      $items = @()
      foreach ($connection in $workbook.Connections) {
        $items += [pscustomobject]@{
          name = [string]$connection.Name
          type = $(try { [int]$connection.Type } catch { $null })
          description = $(try { [string]$connection.Description } catch { '' })
          inModel = $(try { $null -ne $connection.ModelConnection } catch { $false })
        }
      }
      $result.connections = @($items)
    }
    if (Include-ObjectType 'query') {
      $items = @()
      try {
        foreach ($query in $workbook.Queries) {
          $items += [pscustomobject]@{ name = [string]$query.Name; formula = [string]$query.Formula; description = $(try { [string]$query.Description } catch { '' }) }
        }
      } catch {}
      $result.queries = @($items)
    }
    if (Include-ObjectType 'pivotTable') {
      $items = @()
      foreach ($ws in $workbook.Worksheets) {
        foreach ($pivot in $ws.PivotTables()) {
          $items += [pscustomobject]@{
            sheet = [string]$ws.Name
            name = [string]$pivot.Name
            range = $(try { [string]$pivot.TableRange2.Address($false, $false) } catch { '' })
            source = $(try { [string]$pivot.PivotCache().SourceData } catch { '' })
          }
        }
      }
      $result.pivotTables = @($items)
    }
    if (Include-ObjectType 'slicer') {
      $items = @()
      try {
        foreach ($cache in $workbook.SlicerCaches) {
          foreach ($slicer in $cache.Slicers) {
            $items += [pscustomobject]@{
              name = [string]$slicer.Name
              caption = [string]$slicer.Caption
              cache = [string]$cache.Name
              sourceName = $(try { [string]$cache.SourceName } catch { '' })
              sheet = $(try { [string]$slicer.Shape.Parent.Name } catch { '' })
            }
          }
        }
      } catch {}
      $result.slicers = @($items)
    }
    return [pscustomobject]$result
  }
  function Set-PositionAndSize($object) {
    if ($null -ne $actionParams.left) { $object.Left = [double]$actionParams.left }
    if ($null -ne $actionParams.top) { $object.Top = [double]$actionParams.top }
    if ($null -ne $actionParams.width) { $object.Width = [double]$actionParams.width }
    if ($null -ne $actionParams.height) { $object.Height = [double]$actionParams.height }
  }

  if ($_operation -eq 'inspectWorkbookObjects') {
    $operationData.objects = Get-WorkbookObjectSnapshot
    $changes = @()
  } else {
    $objectType = if ($_operation -eq 'manageWorksheetObjects') { 'shape' } else { [string]$actionParams.objectType }
    $command = if ($actionParams.command) { [string]$actionParams.command } elseif ($actionParams.action) { [string]$actionParams.action } else { 'list' }
    if (-not $objectType) { throw 'manageWorkbookObject 需要 params.objectType' }
    $targetSheet = Get-ObjectSheet
    switch ($objectType) {
      'worksheet' {
        switch ($command) {
          'add' {
            $missing = [Type]::Missing
            $newSheet = $workbook.Worksheets.Add($missing, $workbook.Worksheets.Item($workbook.Worksheets.Count), 1, -4167)
            if ($actionParams.name) { $newSheet.Name = [string]$actionParams.name }
          }
          'delete' { $targetSheet.Delete() }
          'rename' { if (-not $actionParams.newName) { throw 'rename 需要 params.newName' }; $targetSheet.Name = [string]$actionParams.newName }
          'copy' {
            $missing = [Type]::Missing
            $after = if ($actionParams.afterSheet) { $workbook.Worksheets.Item([string]$actionParams.afterSheet) } else { $workbook.Worksheets.Item($workbook.Worksheets.Count) }
            $targetSheet.Copy($missing, $after)
            if ($actionParams.newName) { $workbook.ActiveSheet.Name = [string]$actionParams.newName }
          }
          'move' {
            $missing = [Type]::Missing
            if ($actionParams.beforeSheet) { $targetSheet.Move($workbook.Worksheets.Item([string]$actionParams.beforeSheet), $missing) }
            else { $targetSheet.Move($missing, $workbook.Worksheets.Item([string]$actionParams.afterSheet)) }
          }
          'visibility' {
            $targetSheet.Visible = switch ([string]$actionParams.visibility) { 'hidden' { 0 } 'veryHidden' { 2 } default { -1 } }
          }
          'tabColor' { $targetSheet.Tab.Color = Convert-ObjectHexColor $actionParams.color 5263440 }
          'protect' { $targetSheet.Protect([string]$actionParams.password, $true, $true, $true, $true) }
          'unprotect' { $targetSheet.Unprotect([string]$actionParams.password) }
          'list' { }
          default { throw "不支持的工作表命令: $command" }
        }
      }
      'name' {
        $name = [string]$actionParams.name
        if (-not $name) { throw '定义名称操作需要 params.name' }
        $scope = $workbook.Names
        if ($actionParams.scope -and [string]$actionParams.scope -ne 'workbook') { $scope = $workbook.Worksheets.Item([string]$actionParams.scope).Names }
        switch ($command) {
          { $_ -in @('create', 'update', 'upsert') } {
            if (-not $actionParams.refersTo) { throw "$command 需要 params.refersTo" }
            try { $scope.Item($name).Delete() } catch {}
            $definedName = $scope.Add($name, [string]$actionParams.refersTo)
            if ($actionParams.visible -eq $false) { $definedName.Visible = $false }
          }
          'delete' { $scope.Item($name).Delete() }
          'rename' { if (-not $actionParams.newName) { throw 'rename 需要 params.newName' }; $scope.Item($name).Name = [string]$actionParams.newName }
          default { throw "不支持的定义名称命令: $command" }
        }
      }
      'table' {
        $table = $null
        if ($actionParams.name) { try { $table = $targetSheet.ListObjects.Item([string]$actionParams.name) } catch {} }
        switch ($command) {
          'create' {
            $tableRange = $range
            if ($actionParams.range) { $tableRange = $targetSheet.Range([string]$actionParams.range) }
            $missing = [Type]::Missing
            $table = $targetSheet.ListObjects.Add(1, $tableRange, $missing, 1, $missing)
            if ($actionParams.name) { $table.Name = [string]$actionParams.name }
            if ($actionParams.style) { $table.TableStyle = [string]$actionParams.style }
          }
          'delete' { if ($null -eq $table) { throw '找不到结构化表' }; $table.Delete() }
          'unlist' { if ($null -eq $table) { throw '找不到结构化表' }; $table.Unlist() }
          'rename' { if ($null -eq $table -or -not $actionParams.newName) { throw 'rename 需要有效 params.name/newName' }; $table.Name = [string]$actionParams.newName }
          'resize' { if ($null -eq $table -or -not $actionParams.range) { throw 'resize 需要有效 params.name/range' }; $table.Resize($targetSheet.Range([string]$actionParams.range)) }
          'style' { if ($null -eq $table -or -not $actionParams.style) { throw 'style 需要有效 params.name/style' }; $table.TableStyle = [string]$actionParams.style }
          'totals' { if ($null -eq $table) { throw '找不到结构化表' }; $table.ShowTotals = if ($actionParams.enabled -eq $false) { $false } else { $true } }
          'clearFilters' { if ($null -eq $table) { throw '找不到结构化表' }; try { $table.AutoFilter.ShowAllData() } catch {} }
          default { throw "不支持的结构化表命令: $command" }
        }
      }
      'connection' {
        if (-not $actionParams.name) { throw '连接操作需要 params.name' }
        $connection = $workbook.Connections.Item([string]$actionParams.name)
        switch ($command) {
          'refresh' { $connection.Refresh(); try { $app.CalculateUntilAsyncQueriesDone() } catch {} }
          'delete' { $connection.Delete() }
          'rename' { if (-not $actionParams.newName) { throw 'rename 需要 params.newName' }; $connection.Name = [string]$actionParams.newName }
          default { throw "不支持的连接命令: $command" }
        }
      }
      { $_ -in @('shape', 'chart') } {
        if (-not $actionParams.name) { throw "$objectType 操作需要 params.name" }
        $object = if ($objectType -eq 'chart') { $targetSheet.ChartObjects().Item([string]$actionParams.name) } else { $targetSheet.Shapes.Item([string]$actionParams.name) }
        switch ($command) {
          'delete' { $object.Delete() }
          'rename' { if (-not $actionParams.newName) { throw 'rename 需要 params.newName' }; $object.Name = [string]$actionParams.newName }
          { $_ -in @('move', 'resize', 'layout') } { Set-PositionAndSize $object }
          'visibility' { $object.Visible = if ($actionParams.visible -eq $false) { 0 } else { -1 } }
          'duplicate' {
            $copy = $object.Duplicate()
            if ($actionParams.newName) { $copy.Name = [string]$actionParams.newName }
            Set-PositionAndSize $copy
          }
          'alternativeText' { if ($objectType -ne 'shape') { throw 'alternativeText 仅适用于 shape' }; $object.AlternativeText = [string]$actionParams.text }
          'list' { }
          default { throw "不支持的 $objectType 命令: $command" }
        }
      }
      'image' {
        if ($command -ne 'add' -or -not $actionParams.path) { throw 'image 当前需要 command=add 和 params.path' }
        $left = if ($null -ne $actionParams.left) { [double]$actionParams.left } else { [double]$range.Left }
        $top = if ($null -ne $actionParams.top) { [double]$actionParams.top } else { [double]$range.Top }
        $width = if ($null -ne $actionParams.width) { [double]$actionParams.width } else { -1 }
        $height = if ($null -ne $actionParams.height) { [double]$actionParams.height } else { -1 }
        $picture = $targetSheet.Shapes.AddPicture([string]$actionParams.path, $false, $true, $left, $top, $width, $height)
        if ($actionParams.name) { $picture.Name = [string]$actionParams.name }
      }
      'pivotTable' {
        if (-not $actionParams.name) { throw '透视表操作需要 params.name' }
        $pivot = $targetSheet.PivotTables().Item([string]$actionParams.name)
        switch ($command) {
          'refresh' { $pivot.RefreshTable() }
          'clear' { $pivot.TableRange2.Clear() }
          'rename' { if (-not $actionParams.newName) { throw 'rename 需要 params.newName' }; $pivot.Name = [string]$actionParams.newName }
          default { throw "不支持的透视表命令: $command" }
        }
      }
      'slicer' {
        if (-not $actionParams.name) { throw '切片器操作需要 params.name' }
        $slicer = $null; $cache = $null
        foreach ($candidateCache in $workbook.SlicerCaches) {
          foreach ($candidate in $candidateCache.Slicers) {
            if ([string]$candidate.Name -ieq [string]$actionParams.name) { $slicer = $candidate; $cache = $candidateCache; break }
          }
          if ($null -ne $slicer) { break }
        }
        if ($null -eq $slicer) { throw '找不到切片器: ' + [string]$actionParams.name }
        switch ($command) {
          'delete' { $slicer.Delete() }
          'clearFilter' { $cache.ClearManualFilter() }
          'style' { if (-not $actionParams.style) { throw 'style 需要 params.style' }; $slicer.Style = [string]$actionParams.style }
          'columns' { $slicer.NumberOfColumns = [int]$actionParams.columns }
          { $_ -in @('move', 'resize', 'layout') } { Set-PositionAndSize $slicer.Shape }
          default { throw "不支持的切片器命令: $command" }
        }
      }
      default { throw "不支持的工作簿对象类型: $objectType" }
    }
    $operationData.objectType = $objectType
    $operationData.command = $command
    $operationData.objects = Get-WorkbookObjectSnapshot
    $changes += [pscustomobject]@{ kind = 'workbook-object'; target = $objectType; detail = '已执行 ' + $objectType + '/' + $command }
  }
`;
}
