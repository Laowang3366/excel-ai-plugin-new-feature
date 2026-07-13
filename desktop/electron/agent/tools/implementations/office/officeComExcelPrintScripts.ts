const PRINT_OPERATIONS = new Set([
  "inspectPrintSettings",
  "configurePrint",
  "exportSheetsToPdf",
]);

export function isExcelPrintOperation(operation: string): boolean {
  return PRINT_OPERATIONS.has(operation);
}

export function buildExcelPrintOperationScript(operation: string): string | undefined {
  if (!isExcelPrintOperation(operation)) return undefined;
  return `
  function Get-TargetPrintSheets() {
    $requested = @($actionParams.sheetNames | Where-Object { $_ })
    $targets = @()
    if ($requested.Count -gt 0) {
      foreach ($name in $requested) {
        try { $targets += $workbook.Worksheets.Item([string]$name) } catch { throw '找不到导出或打印工作表: ' + [string]$name }
      }
    } elseif ($actionParams.allSheets -eq $true -or $_operation -in @('inspectPrintSettings', 'exportSheetsToPdf')) {
      foreach ($ws in $workbook.Worksheets) { $targets += $ws }
    } else {
      $targets += $sheet
    }
    foreach ($target in $targets) { Write-Output -NoEnumerate $target }
  }
  function Convert-PrintLength([object]$value, [string]$unit) {
    if ($null -eq $value) { return $null }
    if ($unit -eq 'points') { return [double]$value }
    if ($unit -eq 'inches') { return $app.InchesToPoints([double]$value) }
    return $app.CentimetersToPoints([double]$value)
  }
  function Convert-PaperSize([string]$value) {
    switch ($value.ToUpperInvariant()) {
      'LETTER' { 1 }
      'LEGAL' { 5 }
      'A3' { 8 }
      'A4' { 9 }
      'A5' { 11 }
      'B4' { 12 }
      'B5' { 13 }
      'TABLOID' { 3 }
      default { 9 }
    }
  }
  function Set-WorksheetPrintSettings($ws, $config) {
    $setup = $ws.PageSetup
    if ($null -ne $config.printArea) { $setup.PrintArea = [string]$config.printArea }
    if ($config.orientation) { $setup.Orientation = if ([string]$config.orientation -eq 'landscape') { 2 } else { 1 } }
    if ($config.paperSize) { $setup.PaperSize = Convert-PaperSize ([string]$config.paperSize) }
    $marginUnit = if ($config.marginUnit) { [string]$config.marginUnit } else { 'centimeters' }
    if ($config.margins) {
      if ($null -ne $config.margins.top) { $setup.TopMargin = Convert-PrintLength $config.margins.top $marginUnit }
      if ($null -ne $config.margins.bottom) { $setup.BottomMargin = Convert-PrintLength $config.margins.bottom $marginUnit }
      if ($null -ne $config.margins.left) { $setup.LeftMargin = Convert-PrintLength $config.margins.left $marginUnit }
      if ($null -ne $config.margins.right) { $setup.RightMargin = Convert-PrintLength $config.margins.right $marginUnit }
      if ($null -ne $config.margins.header) { $setup.HeaderMargin = Convert-PrintLength $config.margins.header $marginUnit }
      if ($null -ne $config.margins.footer) { $setup.FooterMargin = Convert-PrintLength $config.margins.footer $marginUnit }
    }
    if ($null -ne $config.repeatRows) { $setup.PrintTitleRows = [string]$config.repeatRows }
    if ($null -ne $config.repeatColumns) { $setup.PrintTitleColumns = [string]$config.repeatColumns }
    if ($config.fitToOnePageWide -eq $true) {
      $setup.Zoom = $false
      $setup.FitToPagesWide = 1
      $setup.FitToPagesTall = if ($config.fitToOnePageTall -eq $true) { 1 } else { $false }
    } elseif ($null -ne $config.scale) {
      $setup.Zoom = [Math]::Max(10, [Math]::Min(400, [int]$config.scale))
    } else {
      if ($null -ne $config.fitToPagesWide) { $setup.Zoom = $false; $setup.FitToPagesWide = [int]$config.fitToPagesWide }
      if ($null -ne $config.fitToPagesTall) { $setup.Zoom = $false; $setup.FitToPagesTall = [int]$config.fitToPagesTall }
    }
    if ($null -ne $config.centerHorizontally) { $setup.CenterHorizontally = [bool]$config.centerHorizontally }
    if ($null -ne $config.centerVertically) { $setup.CenterVertically = [bool]$config.centerVertically }
    if ($null -ne $config.printGridlines) { $setup.PrintGridlines = [bool]$config.printGridlines }
    if ($null -ne $config.printHeadings) { $setup.PrintHeadings = [bool]$config.printHeadings }
    if ($null -ne $config.blackAndWhite) { $setup.BlackAndWhite = [bool]$config.blackAndWhite }
    if ($null -ne $config.draft) { $setup.Draft = [bool]$config.draft }
    if ($config.pageOrder) { $setup.Order = if ([string]$config.pageOrder -eq 'overThenDown') { 2 } else { 1 } }
    if ($null -ne $config.firstPageNumber) { $setup.FirstPageNumber = [int]$config.firstPageNumber }
    if ($config.headers) {
      if ($null -ne $config.headers.left) { $setup.LeftHeader = [string]$config.headers.left }
      if ($null -ne $config.headers.center) { $setup.CenterHeader = [string]$config.headers.center }
      if ($null -ne $config.headers.right) { $setup.RightHeader = [string]$config.headers.right }
    }
    if ($config.footers) {
      if ($null -ne $config.footers.left) { $setup.LeftFooter = [string]$config.footers.left }
      if ($null -ne $config.footers.center) { $setup.CenterFooter = [string]$config.footers.center }
      if ($null -ne $config.footers.right) { $setup.RightFooter = [string]$config.footers.right }
    }
    if ($config.clearPageBreaks -eq $true) { $ws.ResetAllPageBreaks() }
    foreach ($address in @($config.horizontalPageBreaks | Where-Object { $_ })) {
      [void]$ws.HPageBreaks.Add($ws.Range([string]$address))
    }
    foreach ($address in @($config.verticalPageBreaks | Where-Object { $_ })) {
      [void]$ws.VPageBreaks.Add($ws.Range([string]$address))
    }
  }
  function Get-WorksheetPrintSnapshot($ws) {
    $setup = $ws.PageSetup
    $horizontalBreaks = @()
    foreach ($break in $ws.HPageBreaks) { $horizontalBreaks += $(try { [string]$break.Location.Address($false, $false) } catch { '' }) }
    $verticalBreaks = @()
    foreach ($break in $ws.VPageBreaks) { $verticalBreaks += $(try { [string]$break.Location.Address($false, $false) } catch { '' }) }
    return [pscustomobject]@{
      sheet = [string]$ws.Name
      printArea = [string]$setup.PrintArea
      orientation = $(if ([int]$setup.Orientation -eq 2) { 'landscape' } else { 'portrait' })
      paperSize = [int]$setup.PaperSize
      marginsPoints = [pscustomobject]@{
        top = [double]$setup.TopMargin
        bottom = [double]$setup.BottomMargin
        left = [double]$setup.LeftMargin
        right = [double]$setup.RightMargin
        header = [double]$setup.HeaderMargin
        footer = [double]$setup.FooterMargin
      }
      repeatRows = [string]$setup.PrintTitleRows
      repeatColumns = [string]$setup.PrintTitleColumns
      zoom = $(try { $setup.Zoom } catch { $null })
      fitToPagesWide = $(try { $setup.FitToPagesWide } catch { $null })
      fitToPagesTall = $(try { $setup.FitToPagesTall } catch { $null })
      centerHorizontally = [bool]$setup.CenterHorizontally
      centerVertically = [bool]$setup.CenterVertically
      printGridlines = [bool]$setup.PrintGridlines
      printHeadings = [bool]$setup.PrintHeadings
      headers = [pscustomobject]@{ left = [string]$setup.LeftHeader; center = [string]$setup.CenterHeader; right = [string]$setup.RightHeader }
      footers = [pscustomobject]@{ left = [string]$setup.LeftFooter; center = [string]$setup.CenterFooter; right = [string]$setup.RightFooter }
      horizontalPageBreaks = @($horizontalBreaks | Where-Object { $_ })
      verticalPageBreaks = @($verticalBreaks | Where-Object { $_ })
    }
  }

  $targetSheets = @(Get-TargetPrintSheets)
  if ($_operation -eq 'inspectPrintSettings') {
    $settings = @()
    foreach ($targetSheet in $targetSheets) { $settings += Get-WorksheetPrintSnapshot $targetSheet }
    $operationData.settings = @($settings)
    $operationData.sheetCount = @($settings).Count
    $changes = @()
  } elseif ($_operation -eq 'configurePrint') {
    $settings = @()
    foreach ($targetSheet in $targetSheets) {
      Set-WorksheetPrintSettings $targetSheet $actionParams
      $settings += Get-WorksheetPrintSnapshot $targetSheet
    }
    $operationData.settings = @($settings)
    $operationData.sheetCount = @($settings).Count
    $changes += [pscustomobject]@{ kind = 'print-settings'; target = @($targetSheets.Name) -join ', '; detail = '已配置页面、打印标题和分页符' }
  } else {
    $exportSettings = if ($actionParams.settings) { $actionParams.settings } else { $null }
    if ($null -ne $exportSettings) { foreach ($targetSheet in $targetSheets) { Set-WorksheetPrintSettings $targetSheet $exportSettings } }
    $mode = if ($actionParams.mode) { [string]$actionParams.mode } else { 'combined' }
    $outputs = @()
    if ($mode -eq 'separate') {
      $outputDirectory = if ($actionParams.outputDirectory) { [string]$actionParams.outputDirectory } else { [IO.Path]::GetDirectoryName($_outputPath) }
      if (-not [IO.Directory]::Exists($outputDirectory)) { [void][IO.Directory]::CreateDirectory($outputDirectory) }
      foreach ($targetSheet in $targetSheets) {
        $safeName = ([string]$targetSheet.Name) -replace '[\\/:*?"<>|]', '_'
        $output = [IO.Path]::Combine($outputDirectory, $safeName + '.pdf')
        if ((Test-Path -LiteralPath $output) -and $actionParams.overwrite -ne $true) { throw 'PDF 已存在，请设置 overwrite=true: ' + $output }
        $targetSheet.ExportAsFixedFormat(0, $output)
        $outputs += $output
      }
      $_outputPath = $outputDirectory
    } else {
      if ((Test-Path -LiteralPath $_outputPath) -and $actionParams.overwrite -ne $true) { throw 'PDF 已存在，请设置 overwrite=true: ' + $_outputPath }
      $first = $true
      foreach ($targetSheet in $targetSheets) { $targetSheet.Select($first); $first = $false }
      $app.ActiveSheet.ExportAsFixedFormat(0, $_outputPath)
      $outputs += $_outputPath
    }
    $operationData.outputPaths = @($outputs)
    $operationData.exportedSheets = @($targetSheets | ForEach-Object { [string]$_.Name })
    $operationData.mode = $mode
    $changes += [pscustomobject]@{ kind = 'export'; target = $_outputPath; detail = '已批量导出指定工作表 PDF' }
  }
`;
}
