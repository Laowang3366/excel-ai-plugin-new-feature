const TEMPLATE_OPERATIONS = new Set([
  "captureWorkbookTemplate",
  "inspectWorkbookFormatting",
  "applyWorkbookTemplate",
]);

export function isWorkbookTemplateOperation(operation: string): boolean {
  return TEMPLATE_OPERATIONS.has(operation);
}

export function buildWorkbookTemplateOperationScript(operation: string): string | undefined {
  if (!isWorkbookTemplateOperation(operation)) return undefined;
  return `
  function Convert-TemplateColor([object]$value, [int]$fallback) {
    if ($null -eq $value) { return $fallback }
    if ($value -is [int] -or $value -is [long] -or $value -is [double]) { return [int]$value }
    $hex = ([string]$value).Trim().TrimStart('#')
    if ($hex -notmatch '^[0-9a-fA-F]{6}$') { return $fallback }
    return [Convert]::ToInt32($hex.Substring(0, 2), 16) +
      (256 * [Convert]::ToInt32($hex.Substring(2, 2), 16)) +
      (65536 * [Convert]::ToInt32($hex.Substring(4, 2), 16))
  }
  function Convert-HorizontalAlignment([string]$value) {
    switch ($value) { 'left' { -4131 } '-4131' { -4131 } 'center' { -4108 } '-4108' { -4108 } 'right' { -4152 } '-4152' { -4152 } 'fill' { 5 } '5' { 5 } 'justify' { -4130 } '-4130' { -4130 } default { 1 } }
  }
  function Convert-VerticalAlignment([string]$value) {
    switch ($value) { 'top' { -4160 } '-4160' { -4160 } 'bottom' { -4107 } '-4107' { -4107 } default { -4108 } }
  }
  function Set-TemplateRangeStyle($targetRange, $style) {
    if ($null -eq $style -or $null -eq $targetRange) { return }
    if ($style.fontName) { $targetRange.Font.Name = [string]$style.fontName }
    if ($null -ne $style.fontSize) { $targetRange.Font.Size = [double]$style.fontSize }
    if ($null -ne $style.bold) { $targetRange.Font.Bold = [bool]$style.bold }
    if ($null -ne $style.italic) { $targetRange.Font.Italic = [bool]$style.italic }
    if ($null -ne $style.underline) { $targetRange.Font.Underline = if ($style.underline -eq $false) { -4142 } else { 2 } }
    if ($null -ne $style.fontColor) { $targetRange.Font.Color = Convert-TemplateColor $style.fontColor 0 }
    if ($null -ne $style.fillColor) { $targetRange.Interior.Color = Convert-TemplateColor $style.fillColor 16777215 }
    if ($null -ne $style.numberFormat) { $targetRange.NumberFormat = [string]$style.numberFormat }
    if ($style.horizontalAlignment) { $targetRange.HorizontalAlignment = Convert-HorizontalAlignment ([string]$style.horizontalAlignment) }
    if ($style.verticalAlignment) { $targetRange.VerticalAlignment = Convert-VerticalAlignment ([string]$style.verticalAlignment) }
    if ($null -ne $style.wrapText) { $targetRange.WrapText = [bool]$style.wrapText }
    if ($null -ne $style.indentLevel) { $targetRange.IndentLevel = [int]$style.indentLevel }
    if ($null -ne $style.rowHeight) { $targetRange.RowHeight = [double]$style.rowHeight }
    if ($null -ne $style.columnWidth) { $targetRange.ColumnWidth = [double]$style.columnWidth }
    if ($style.borders) {
      $lineStyle = if ($style.borders.lineStyle) { [int]$style.borders.lineStyle } else { 1 }
      $weight = if ($style.borders.weight) { [int]$style.borders.weight } else { 2 }
      $color = Convert-TemplateColor $style.borders.color 14277081
      foreach ($edge in @(7, 8, 9, 10, 11, 12)) {
        try { $targetRange.Borders.Item($edge).LineStyle = $lineStyle; $targetRange.Borders.Item($edge).Weight = $weight; $targetRange.Borders.Item($edge).Color = $color } catch {}
      }
    }
  }
  function Get-DefaultProfessionalTemplate([string]$preset) {
    $accent = switch ($preset) { 'financial' { '217346' } 'dashboard' { '202124' } 'minimal' { '5F6368' } default { '1F4E79' } }
    $headerFill = switch ($preset) { 'financial' { '217346' } 'dashboard' { '202124' } 'minimal' { 'E8EAED' } default { '1F4E79' } }
    $headerFont = if ($preset -eq 'minimal') { '202124' } else { 'FFFFFF' }
    $bandFill = switch ($preset) { 'financial' { 'E2F0D9' } 'dashboard' { 'F1F3F4' } 'minimal' { 'F8F9FA' } default { 'D9EAF7' } }
    return [pscustomobject]@{
      version = 1
      preset = $preset
      theme = [pscustomobject]@{ fontName = '微软雅黑'; fontSize = 10.5; fontColor = '202124'; accentColor = $accent; backgroundColor = 'FFFFFF' }
      defaultSheet = [pscustomobject]@{
        titleRows = 0
        headerRows = 1
        totalRows = 0
        showGridlines = $false
        freezeRows = 1
        freezeColumns = 0
        autoFit = $true
        baseStyle = [pscustomobject]@{ verticalAlignment = 'center' }
        headerStyle = [pscustomobject]@{ fillColor = $headerFill; fontColor = $headerFont; bold = $true; horizontalAlignment = 'center'; wrapText = $true; rowHeight = 24; borders = [pscustomobject]@{ color = $accent; weight = 2 } }
        titleStyle = [pscustomobject]@{ fillColor = $accent; fontColor = 'FFFFFF'; bold = $true; fontSize = 16; horizontalAlignment = 'left'; rowHeight = 30 }
        totalStyle = [pscustomobject]@{ fillColor = $bandFill; bold = $true; borders = [pscustomobject]@{ color = $accent; weight = 2 } }
        bandedRows = $true
        bandedRowColor = $bandFill
        tableStyle = $(if ($preset -eq 'financial') { 'TableStyleMedium4' } else { 'TableStyleMedium2' })
      }
      sheets = @()
    }
  }
  function Get-SheetTemplateRule($template, [string]$sheetName, [int]$sheetIndex) {
    $sourceName = $sheetName
    if ($actionParams.sheetMap) {
      $mapped = $actionParams.sheetMap.PSObject.Properties[$sheetName]
      if ($null -ne $mapped) { $sourceName = [string]$mapped.Value }
    }
    foreach ($rule in @($template.sheets)) {
      if ([string]$rule.name -ieq $sourceName -or [string]$rule.targetName -ieq $sheetName) { return $rule }
    }
    if ($actionParams.matchSheetsByIndex -ne $false -and @($template.sheets).Count -ge $sheetIndex) { return @($template.sheets)[$sheetIndex - 1] }
    return $template.defaultSheet
  }
  function Get-RuleValue($rule, $fallback, [string]$name) {
    if ($null -ne $rule -and $null -ne $rule.PSObject.Properties[$name] -and $null -ne $rule.$name) { return $rule.$name }
    if ($null -ne $fallback -and $null -ne $fallback.PSObject.Properties[$name]) { return $fallback.$name }
    return $null
  }
  function Get-ColumnRange($ws, $used, $columnRule) {
    if ($columnRule.range) { $resolved = $ws.Range([string]$columnRule.range); Write-Output -NoEnumerate $resolved; return }
    $index = $null
    if ($columnRule.index) { $index = [int]$columnRule.index }
    elseif ($columnRule.letter) { $index = [int]$ws.Range(([string]$columnRule.letter) + '1').Column }
    elseif ($columnRule.header) {
      for ($column = 1; $column -le $used.Columns.Count; $column++) {
        if ([string]$used.Cells.Item(1, $column).Text -ieq [string]$columnRule.header) { $index = $column; break }
      }
    }
    if ($null -eq $index) { return $null }
    $resolved = $used.Columns.Item($index)
    Write-Output -NoEnumerate $resolved
  }
  function Add-TemplateConditionalFormat($targetRange, $rule) {
    $condition = $null
    switch ([string]$rule.type) {
      'formula' { if (-not $rule.formula) { throw 'formula 条件格式需要 formula' }; $condition = $targetRange.FormatConditions.Add(2, 3, [string]$rule.formula) }
      'cellValue' {
        $operator = switch ([string]$rule.operator) { 'between' { 1 } 'notBetween' { 2 } 'equal' { 3 } 'notEqual' { 4 } 'greater' { 5 } 'less' { 6 } 'greaterOrEqual' { 7 } 'lessOrEqual' { 8 } default { 3 } }
        $condition = $targetRange.FormatConditions.Add(1, $operator, $rule.value1, $rule.value2)
      }
      'dataBar' {
        $condition = $targetRange.FormatConditions.AddDatabar()
        if ($rule.color) { $condition.BarColor.Color = Convert-TemplateColor $rule.color 5263440 }
      }
      'colorScale' {
        $condition = $targetRange.FormatConditions.AddColorScale(3)
        $colors = @($rule.colors)
        if ($colors.Count -ge 3) {
          $condition.ColorScaleCriteria.Item(1).FormatColor.Color = Convert-TemplateColor $colors[0] 13551615
          $condition.ColorScaleCriteria.Item(2).FormatColor.Color = Convert-TemplateColor $colors[1] 10284031
          $condition.ColorScaleCriteria.Item(3).FormatColor.Color = Convert-TemplateColor $colors[2] 5296274
        }
      }
      'iconSet' {
        $condition = $targetRange.FormatConditions.AddIconSetCondition()
        $condition.IconSet = $workbook.IconSets.Item($(if ($rule.iconSetIndex) { [int]$rule.iconSetIndex } else { 1 }))
      }
      default { throw '不支持的条件格式类型: ' + [string]$rule.type }
    }
    if ($null -ne $condition -and $rule.type -in @('formula', 'cellValue')) {
      if ($rule.fillColor) { $condition.Interior.Color = Convert-TemplateColor $rule.fillColor 16777164 }
      if ($rule.fontColor) { $condition.Font.Color = Convert-TemplateColor $rule.fontColor 0 }
      if ($null -ne $rule.bold) { $condition.Font.Bold = [bool]$rule.bold }
    }
  }
  function Capture-WorkbookTemplate() {
    $sheetRules = @()
    foreach ($ws in $workbook.Worksheets) {
      $ws.Activate()
      $used = $ws.UsedRange
      $columns = @()
      for ($column = 1; $column -le $used.Columns.Count; $column++) {
        $columnRange = $used.Columns.Item($column)
        $columns += [pscustomobject]@{
          index = $column
          header = [string]$used.Cells.Item(1, $column).Text
          width = $(try { [double]$columnRange.ColumnWidth } catch { $null })
          numberFormat = $(try { [string]$used.Cells.Item([Math]::Min(2, $used.Rows.Count), $column).NumberFormat } catch { 'General' })
          horizontalAlignment = $(try { [int]$columnRange.HorizontalAlignment } catch { $null })
          hidden = $(try { [bool]$columnRange.EntireColumn.Hidden } catch { $false })
        }
      }
      $header = $used.Rows.Item(1)
      $sheetRules += [pscustomobject]@{
        name = [string]$ws.Name
        headerRows = if ($actionParams.headerRows) { [int]$actionParams.headerRows } else { 1 }
        titleRows = if ($actionParams.titleRows) { [int]$actionParams.titleRows } else { 0 }
        totalRows = if ($actionParams.totalRows) { [int]$actionParams.totalRows } else { 0 }
        showGridlines = $(try { [bool]$app.ActiveWindow.DisplayGridlines } catch { $true })
        freezeRows = $(try { if ($app.ActiveWindow.FreezePanes) { [int]$app.ActiveWindow.SplitRow } else { 0 } } catch { 0 })
        freezeColumns = $(try { if ($app.ActiveWindow.FreezePanes) { [int]$app.ActiveWindow.SplitColumn } else { 0 } } catch { 0 })
        baseStyle = [pscustomobject]@{
          fontName = $(try { [string]$used.Font.Name } catch { '微软雅黑' })
          fontSize = $(try { [double]$used.Font.Size } catch { 10.5 })
          fontColor = $(try { [int]$used.Font.Color } catch { 0 })
          verticalAlignment = $(try { [int]$used.VerticalAlignment } catch { $null })
          wrapText = $(try { [bool]$used.WrapText } catch { $false })
        }
        headerStyle = [pscustomobject]@{
          fillColor = $(try { [int]$header.Interior.Color } catch { 16777215 })
          fontColor = $(try { [int]$header.Font.Color } catch { 0 })
          bold = $(try { [bool]$header.Font.Bold } catch { $false })
          fontSize = $(try { [double]$header.Font.Size } catch { 10.5 })
          rowHeight = $(try { [double]$header.RowHeight } catch { $null })
          horizontalAlignment = $(try { [int]$header.HorizontalAlignment } catch { $null })
        }
        columns = @($columns)
        tableStyle = $(try { $(if ($ws.ListObjects.Count -gt 0) { [string]$ws.ListObjects.Item(1).TableStyle } else { '' }) } catch { '' })
        print = [pscustomobject]@{
          area = $(try { [string]$ws.PageSetup.PrintArea } catch { '' })
          orientation = $(try { $(if ([int]$ws.PageSetup.Orientation -eq 2) { 'landscape' } else { 'portrait' }) } catch { 'portrait' })
          fitToPagesWide = $(try { [int]$ws.PageSetup.FitToPagesWide } catch { 1 })
          fitToPagesTall = $(try { [int]$ws.PageSetup.FitToPagesTall } catch { 1 })
          repeatRows = $(try { [string]$ws.PageSetup.PrintTitleRows } catch { '' })
          header = $(try { [string]$ws.PageSetup.CenterHeader } catch { '' })
          footer = $(try { [string]$ws.PageSetup.CenterFooter } catch { '' })
        }
      }
    }
    $firstUsed = $workbook.Worksheets.Item(1).UsedRange
    return [pscustomobject]@{
      version = 1
      capturedFrom = [string]$workbook.Name
      capturedAt = [DateTime]::UtcNow.ToString('o')
      theme = [pscustomobject]@{
        fontName = $(try { [string]$firstUsed.Font.Name } catch { '微软雅黑' })
        fontSize = $(try { [double]$firstUsed.Font.Size } catch { 10.5 })
        fontColor = $(try { [int]$firstUsed.Font.Color } catch { 0 })
      }
      sheets = @($sheetRules)
    }
  }

  if ($_operation -in @('captureWorkbookTemplate', 'inspectWorkbookFormatting')) {
    $operationData.template = Capture-WorkbookTemplate
    $operationData.sheetCount = @($operationData.template.sheets).Count
    $changes = @()
  } else {
    $preset = if ($actionParams.preset) { [string]$actionParams.preset } else { 'professional' }
    $template = if ($null -ne $actionParams.template) { $actionParams.template } else { Get-DefaultProfessionalTemplate $preset }
    $defaultRule = $template.defaultSheet
    $targetNames = @($actionParams.sheetNames | Where-Object { $_ })
    $appliedSheets = @()
    $rulesApplied = 0
    foreach ($ws in $workbook.Worksheets) {
      if ($targetNames.Count -gt 0 -and $targetNames -notcontains [string]$ws.Name) { continue }
      if ($actionParams.allSheets -eq $false -and [string]$ws.Name -ine $_sheetName) { continue }
      $used = $ws.UsedRange
      if ($used.Cells.Count -eq 1 -and [string]::IsNullOrWhiteSpace([string]$used.Text)) { continue }
      $rule = Get-SheetTemplateRule $template ([string]$ws.Name) ([int]$ws.Index)
      if ($null -eq $rule) { $rule = $defaultRule }
      $theme = $template.theme
      if ($theme.fontName) { $used.Font.Name = [string]$theme.fontName }
      if ($null -ne $theme.fontSize) { $used.Font.Size = [double]$theme.fontSize }
      if ($null -ne $theme.fontColor) { $used.Font.Color = Convert-TemplateColor $theme.fontColor 0 }
      Set-TemplateRangeStyle $used (Get-RuleValue $rule $defaultRule 'baseStyle')
      $titleRows = [int](Get-RuleValue $rule $defaultRule 'titleRows')
      $headerRows = [int](Get-RuleValue $rule $defaultRule 'headerRows')
      $totalRows = [int](Get-RuleValue $rule $defaultRule 'totalRows')
      if ($titleRows -gt 0) {
        $titleRange = $used.Rows.Item('1:' + [Math]::Min($titleRows, $used.Rows.Count))
        Set-TemplateRangeStyle $titleRange (Get-RuleValue $rule $defaultRule 'titleStyle')
      }
      if ($headerRows -gt 0 -and $used.Rows.Count -gt $titleRows) {
        $headerStart = $titleRows + 1
        $headerEnd = [Math]::Min($used.Rows.Count, $titleRows + $headerRows)
        $headerRange = $used.Rows.Item($headerStart.ToString() + ':' + $headerEnd.ToString())
        Set-TemplateRangeStyle $headerRange (Get-RuleValue $rule $defaultRule 'headerStyle')
      }
      if ($totalRows -gt 0) {
        $totalStart = [Math]::Max(1, $used.Rows.Count - $totalRows + 1)
        $totalRange = $used.Rows.Item($totalStart.ToString() + ':' + $used.Rows.Count.ToString())
        Set-TemplateRangeStyle $totalRange (Get-RuleValue $rule $defaultRule 'totalStyle')
      }
      $bandedRows = Get-RuleValue $rule $defaultRule 'bandedRows'
      if ($bandedRows -eq $true -and $used.Rows.Count -gt ($titleRows + $headerRows + $totalRows)) {
        $bodyStart = $titleRows + $headerRows + 1
        $bodyEnd = $used.Rows.Count - $totalRows
        $bodyRange = $used.Rows.Item($bodyStart.ToString() + ':' + $bodyEnd.ToString())
        $bandColor = Convert-TemplateColor (Get-RuleValue $rule $defaultRule 'bandedRowColor') 16119285
        $formula = '=MOD(ROW()-' + $bodyStart + ',2)=1'
        $condition = $bodyRange.FormatConditions.Add(2, 3, $formula)
        $condition.Interior.Color = $bandColor
      }
      foreach ($columnRule in @($rule.columns)) {
        $columnRange = Get-ColumnRange $ws $used $columnRule
        if ($null -eq $columnRange) { continue }
        if ($null -ne $columnRule.width) { $columnRange.ColumnWidth = [double]$columnRule.width }
        if ($columnRule.autoFit -eq $true) { [void]$columnRange.AutoFit() }
        if ($columnRule.numberFormat) { $columnRange.NumberFormat = [string]$columnRule.numberFormat }
        if ($columnRule.horizontalAlignment) { $columnRange.HorizontalAlignment = Convert-HorizontalAlignment ([string]$columnRule.horizontalAlignment) }
        if ($null -ne $columnRule.hidden) { $columnRange.EntireColumn.Hidden = [bool]$columnRule.hidden }
        if ($columnRule.style) { Set-TemplateRangeStyle $columnRange $columnRule.style }
        $rulesApplied++
      }
      foreach ($rangeRule in @($rule.ranges)) {
        if (-not $rangeRule.range) { continue }
        $targetRange = $ws.Range([string]$rangeRule.range)
        if ($rangeRule.merge -eq $true) { $targetRange.Merge() }
        Set-TemplateRangeStyle $targetRange $rangeRule.style
        $rulesApplied++
      }
      foreach ($conditionalRule in @($rule.conditionalFormats)) {
        if (-not $conditionalRule.range) { continue }
        $targetRange = $ws.Range([string]$conditionalRule.range)
        if ($conditionalRule.replace -eq $true) { $targetRange.FormatConditions.Delete() }
        Add-TemplateConditionalFormat $targetRange $conditionalRule
        $rulesApplied++
      }
      $tableStyle = Get-RuleValue $rule $defaultRule 'tableStyle'
      if ($tableStyle) { foreach ($table in $ws.ListObjects) { $table.TableStyle = [string]$tableStyle } }
      if ($rule.createTable -eq $true -and $ws.ListObjects.Count -eq 0) {
        $missing = [Type]::Missing
        $table = $ws.ListObjects.Add(1, $used, $missing, 1, $missing)
        if ($rule.tableName) { $table.Name = [string]$rule.tableName }
        if ($tableStyle) { $table.TableStyle = [string]$tableStyle }
      }
      if ((Get-RuleValue $rule $defaultRule 'autoFit') -ne $false) { [void]$used.Columns.AutoFit(); [void]$used.Rows.AutoFit() }
      $maxColumnWidth = Get-RuleValue $rule $defaultRule 'maxColumnWidth'
      if ($maxColumnWidth) { foreach ($column in $used.Columns) { if ($column.ColumnWidth -gt [double]$maxColumnWidth) { $column.ColumnWidth = [double]$maxColumnWidth } } }
      $freezeRows = [int](Get-RuleValue $rule $defaultRule 'freezeRows')
      $freezeColumns = [int](Get-RuleValue $rule $defaultRule 'freezeColumns')
      $ws.Activate()
      $app.ActiveWindow.FreezePanes = $false
      if ($freezeRows -gt 0 -or $freezeColumns -gt 0) { $ws.Cells.Item($freezeRows + 1, $freezeColumns + 1).Select(); $app.ActiveWindow.FreezePanes = $true }
      $showGridlines = Get-RuleValue $rule $defaultRule 'showGridlines'
      if ($null -ne $showGridlines) { $app.ActiveWindow.DisplayGridlines = [bool]$showGridlines }
      if ($null -ne $rule.tabColor) { $ws.Tab.Color = Convert-TemplateColor $rule.tabColor 5263440 }
      if ($rule.print) {
        $setup = $ws.PageSetup
        if ($rule.print.area) { $setup.PrintArea = [string]$rule.print.area }
        if ($rule.print.orientation) { $setup.Orientation = if ([string]$rule.print.orientation -eq 'landscape') { 2 } else { 1 } }
        if ($null -ne $rule.print.fitToPagesWide) { $setup.Zoom = $false; $setup.FitToPagesWide = [int]$rule.print.fitToPagesWide }
        if ($null -ne $rule.print.fitToPagesTall) { $setup.Zoom = $false; $setup.FitToPagesTall = [int]$rule.print.fitToPagesTall }
        if ($rule.print.repeatRows) { $setup.PrintTitleRows = [string]$rule.print.repeatRows }
        if ($rule.print.header) { $setup.CenterHeader = [string]$rule.print.header }
        if ($rule.print.footer) { $setup.CenterFooter = [string]$rule.print.footer }
      }
      $appliedSheets += [pscustomobject]@{
        name = [string]$ws.Name
        range = [string]$used.Address($false, $false)
        rows = [int]$used.Rows.Count
        columns = [int]$used.Columns.Count
        tables = [int]$ws.ListObjects.Count
      }
    }
    $operationData.preset = $preset
    $operationData.appliedSheets = @($appliedSheets)
    $operationData.appliedSheetCount = @($appliedSheets).Count
    $operationData.rulesApplied = $rulesApplied
    $changes += [pscustomobject]@{ kind = 'workbook-template'; target = $preset; detail = '已应用专业工作簿模板到 ' + @($appliedSheets).Count + ' 个工作表' }
  }
`;
}
