const CHART_OPERATIONS = new Set(["inspectCharts", "formatChart"]);

export function isChartAdvancedOperation(operation: string): boolean {
  return CHART_OPERATIONS.has(operation);
}

export function buildChartAdvancedOperationScript(operation: string): string | undefined {
  if (!isChartAdvancedOperation(operation)) return undefined;
  return `
  function Convert-ExcelHexColor([object]$value, [int]$fallback) {
    if ($null -eq $value) { return $fallback }
    $hex = ([string]$value).Trim().TrimStart('#')
    if ($hex -notmatch '^[0-9a-fA-F]{6}$') { return $fallback }
    $r = [Convert]::ToInt32($hex.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($hex.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($hex.Substring(4, 2), 16)
    return $r + (256 * $g) + (65536 * $b)
  }
  function Convert-ExcelChartType([string]$name) {
    switch ($name.ToLowerInvariant()) {
      'column' { 51 }
      'columnclustered' { 51 }
      'columnstacked' { 52 }
      'columnstacked100' { 53 }
      'bar' { 57 }
      'barclustered' { 57 }
      'barstacked' { 58 }
      'line' { 4 }
      'linemarkers' { 65 }
      'pie' { 5 }
      'doughnut' { -4120 }
      'area' { 1 }
      'areastacked' { 76 }
      'scatter' { -4169 }
      'scatterlines' { 74 }
      'bubble' { 15 }
      'radar' { -4151 }
      'stock' { 88 }
      default { 51 }
    }
  }
  function Resolve-ExcelRange([string]$reference) {
    if (-not $reference) { return $null }
    $parts = $reference -split '!', 2
    if ($parts.Count -eq 2) { $resolved = $workbook.Worksheets.Item($parts[0].Trim("'")).Range($parts[1]) }
    else { $resolved = $sheet.Range($reference) }
    Write-Output -NoEnumerate $resolved
  }
  function Get-TargetChartObject() {
    if ($actionParams.chartName) {
      foreach ($ws in $workbook.Worksheets) {
        try { return $ws.ChartObjects().Item([string]$actionParams.chartName) } catch {}
      }
      throw '找不到图表: ' + [string]$actionParams.chartName
    }
    $index = if ($actionParams.chartIndex) { [int]$actionParams.chartIndex } else { 1 }
    if ($sheet.ChartObjects().Count -lt $index) { throw "找不到图表: $index" }
    return $sheet.ChartObjects().Item($index)
  }
  function Get-AxisSnapshot($chart, [int]$axisType, [int]$axisGroup, [string]$kind) {
    try {
      $axis = $chart.Axes($axisType, $axisGroup)
      return [pscustomobject]@{
        kind = $kind
        group = $(if ($axisGroup -eq 2) { 'secondary' } else { 'primary' })
        title = $(if ($axis.HasTitle) { [string]$axis.AxisTitle.Text } else { '' })
        minimum = $(if ($axis.MinimumScaleIsAuto) { $null } else { [double]$axis.MinimumScale })
        maximum = $(if ($axis.MaximumScaleIsAuto) { $null } else { [double]$axis.MaximumScale })
        majorUnit = $(try { [double]$axis.MajorUnit } catch { $null })
        minorUnit = $(try { [double]$axis.MinorUnit } catch { $null })
        numberFormat = $(try { [string]$axis.TickLabels.NumberFormat } catch { '' })
        reverse = $(try { [bool]$axis.ReversePlotOrder } catch { $false })
      }
    } catch { return $null }
  }
  function Get-ChartSnapshot($chartObject, [string]$worksheetName) {
    $chart = $chartObject.Chart
    $series = @()
    for ($index = 1; $index -le $chart.SeriesCollection().Count; $index++) {
      $item = $chart.SeriesCollection($index)
      $series += [pscustomobject]@{
        index = $index
        name = $(try { [string]$item.Name } catch { '' })
        formula = $(try { [string]$item.Formula } catch { '' })
        chartType = $(try { [int]$item.ChartType } catch { [int]$chart.ChartType })
        axisGroup = $(try { $(if ([int]$item.AxisGroup -eq 2) { 'secondary' } else { 'primary' }) } catch { 'primary' })
        hasDataLabels = $(try { [bool]$item.HasDataLabels } catch { $false })
      }
    }
    $axes = @(
      Get-AxisSnapshot $chart 1 1 'category'
      Get-AxisSnapshot $chart 2 1 'value'
      Get-AxisSnapshot $chart 1 2 'category'
      Get-AxisSnapshot $chart 2 2 'value'
    ) | Where-Object { $null -ne $_ }
    return [pscustomobject]@{
      sheet = $worksheetName
      name = [string]$chartObject.Name
      title = $(if ($chart.HasTitle) { [string]$chart.ChartTitle.Text } else { '' })
      chartType = [int]$chart.ChartType
      style = $(try { [int]$chart.ChartStyle } catch { $null })
      legend = [pscustomobject]@{
        visible = [bool]$chart.HasLegend
        position = $(try { [int]$chart.Legend.Position } catch { $null })
      }
      position = [pscustomobject]@{
        left = [double]$chartObject.Left
        top = [double]$chartObject.Top
        width = [double]$chartObject.Width
        height = [double]$chartObject.Height
      }
      series = @($series)
      axes = @($axes)
    }
  }
  function Set-ChartAxis($chart, $config) {
    $axisType = if ([string]$config.kind -eq 'category') { 1 } else { 2 }
    $axisGroup = if ([string]$config.group -eq 'secondary') { 2 } else { 1 }
    try { $axis = $chart.Axes($axisType, $axisGroup) } catch {
      if ($null -ne $config.visible -and $config.visible -eq $false) { return }
      throw '图表不存在指定坐标轴: ' + [string]$config.kind + '/' + [string]$config.group
    }
    if ($null -ne $config.visible -and $config.visible -eq $false) { $axis.Delete(); return }
    if ($null -ne $config.title) {
      $axis.HasTitle = -not [string]::IsNullOrWhiteSpace([string]$config.title)
      if ($axis.HasTitle) { $axis.AxisTitle.Text = [string]$config.title }
    }
    if ($null -ne $config.minimum) { $axis.MinimumScaleIsAuto = $false; $axis.MinimumScale = [double]$config.minimum }
    elseif ($config.autoMinimum -eq $true) { $axis.MinimumScaleIsAuto = $true }
    if ($null -ne $config.maximum) { $axis.MaximumScaleIsAuto = $false; $axis.MaximumScale = [double]$config.maximum }
    elseif ($config.autoMaximum -eq $true) { $axis.MaximumScaleIsAuto = $true }
    if ($null -ne $config.majorUnit) { $axis.MajorUnitIsAuto = $false; $axis.MajorUnit = [double]$config.majorUnit }
    if ($null -ne $config.minorUnit) { $axis.MinorUnitIsAuto = $false; $axis.MinorUnit = [double]$config.minorUnit }
    if ($config.numberFormat) { $axis.TickLabels.NumberFormat = [string]$config.numberFormat }
    if ($null -ne $config.reverse) { $axis.ReversePlotOrder = [bool]$config.reverse }
    if ($null -ne $config.crossesAt) { $axis.CrossesAt = [double]$config.crossesAt }
    if ($null -ne $config.logBase) { $axis.ScaleType = -4133; $axis.LogBase = [double]$config.logBase }
    elseif ($config.linear -eq $true) { $axis.ScaleType = -4132 }
    if ($null -ne $config.majorGridlines) { $axis.HasMajorGridlines = [bool]$config.majorGridlines }
    if ($config.gridlineColor -and $axis.HasMajorGridlines) {
      $axis.MajorGridlines.Format.Line.ForeColor.RGB = Convert-ExcelHexColor $config.gridlineColor 13882323
    }
  }
  function Set-SeriesDataLabels($series, $config) {
    if ($config.enabled -eq $false) { try { $series.DataLabels().Delete() } catch {}; return }
    $series.ApplyDataLabels()
    $labels = $series.DataLabels()
    if ($null -ne $config.showValue) { $labels.ShowValue = [bool]$config.showValue }
    if ($null -ne $config.showCategoryName) { $labels.ShowCategoryName = [bool]$config.showCategoryName }
    if ($null -ne $config.showSeriesName) { $labels.ShowSeriesName = [bool]$config.showSeriesName }
    if ($null -ne $config.showPercentage) { $labels.ShowPercentage = [bool]$config.showPercentage }
    if ($config.numberFormat) { $labels.NumberFormat = [string]$config.numberFormat }
    if ($config.position) {
      $labels.Position = switch ([string]$config.position) {
        'center' { -4108 }
        'insideEnd' { 3 }
        'insideBase' { 4 }
        'outsideEnd' { 2 }
        'above' { 0 }
        'below' { 1 }
        default { -4108 }
      }
    }
  }
  function Set-ChartSeries($chart, $config) {
    $series = $null
    if ($config.index) { try { $series = $chart.SeriesCollection([int]$config.index) } catch {} }
    if ($null -eq $series -and $config.matchName) {
      foreach ($candidate in $chart.SeriesCollection()) { if ([string]$candidate.Name -ieq [string]$config.matchName) { $series = $candidate; break } }
    }
    $command = if ($config.command) { [string]$config.command } else { 'update' }
    if ($command -eq 'add') { $series = $chart.SeriesCollection().NewSeries() }
    if ($null -eq $series) { throw '找不到要编辑的数据系列' }
    if ($command -eq 'delete') { $series.Delete(); return }
    if ($config.formula) { $series.Formula = [string]$config.formula }
    if ($null -ne $config.name) { $series.Name = [string]$config.name }
    if ($null -ne $config.values) {
      if ($config.values -is [string]) { $series.Values = Resolve-ExcelRange ([string]$config.values) } else { $series.Values = @($config.values) }
    }
    $categoryValues = if ($null -ne $config.categories) { $config.categories } else { $config.xValues }
    if ($null -ne $categoryValues) {
      if ($categoryValues -is [string]) { $series.XValues = Resolve-ExcelRange ([string]$categoryValues) } else { $series.XValues = @($categoryValues) }
    }
    if ($config.chartType) { $series.ChartType = Convert-ExcelChartType ([string]$config.chartType) }
    if ($config.axisGroup) { $series.AxisGroup = if ([string]$config.axisGroup -eq 'secondary') { 2 } else { 1 } }
    if ($config.markerStyle) {
      $series.MarkerStyle = switch ([string]$config.markerStyle) { 'circle' { 8 } 'square' { 1 } 'diamond' { 2 } 'triangle' { 3 } 'none' { -4142 } default { 8 } }
    }
    if ($null -ne $config.markerSize) { $series.MarkerSize = [int]$config.markerSize }
    if ($config.fillColor) {
      $series.Format.Fill.Visible = -1
      $series.Format.Fill.Solid()
      $series.Format.Fill.ForeColor.RGB = Convert-ExcelHexColor $config.fillColor 5263440
      if ($null -ne $config.fillTransparency) { $series.Format.Fill.Transparency = [double]$config.fillTransparency }
    }
    if ($config.lineColor -or $null -ne $config.lineVisible) {
      $series.Format.Line.Visible = if ($config.lineVisible -eq $false) { 0 } else { -1 }
      if ($config.lineColor) { $series.Format.Line.ForeColor.RGB = Convert-ExcelHexColor $config.lineColor 5263440 }
      if ($null -ne $config.lineWeight) { $series.Format.Line.Weight = [double]$config.lineWeight }
    }
    if ($null -ne $config.smooth) { try { $series.Smooth = [bool]$config.smooth } catch {} }
    if ($config.dataLabels) { Set-SeriesDataLabels $series $config.dataLabels }
    if ($config.trendline) {
      if ($config.trendline.replace -ne $false) { try { while ($series.Trendlines().Count -gt 0) { $series.Trendlines(1).Delete() } } catch {} }
      $trendType = switch ([string]$config.trendline.type) { 'exponential' { 5 } 'logarithmic' { -4133 } 'polynomial' { 3 } 'power' { 4 } 'movingAverage' { 6 } default { -4132 } }
      $trend = $series.Trendlines().Add($trendType)
      if ($config.trendline.name) { $trend.Name = [string]$config.trendline.name }
      if ($config.trendline.order) { try { $trend.Order = [int]$config.trendline.order } catch {} }
      if ($config.trendline.period) { try { $trend.Period = [int]$config.trendline.period } catch {} }
      if ($null -ne $config.trendline.displayEquation) { $trend.DisplayEquation = [bool]$config.trendline.displayEquation }
      if ($null -ne $config.trendline.displayRSquared) { $trend.DisplayRSquared = [bool]$config.trendline.displayRSquared }
    }
    if ($config.errorBars) {
      $amount = if ($config.errorBars.amount) { [double]$config.errorBars.amount } else { 1 }
      $series.ErrorBar(1, 1, 2, $amount)
    }
  }

  if ($_operation -eq 'inspectCharts') {
    $charts = @()
    foreach ($ws in $workbook.Worksheets) {
      foreach ($chartObject in $ws.ChartObjects()) {
        if ($actionParams.chartName -and [string]$chartObject.Name -ine [string]$actionParams.chartName) { continue }
        $charts += Get-ChartSnapshot $chartObject ([string]$ws.Name)
      }
    }
    $operationData.charts = @($charts)
    $operationData.chartCount = @($charts).Count
    $changes = @()
  } else {
    $chartObject = Get-TargetChartObject
    $chart = $chartObject.Chart
    if ($actionParams.sourceRange) { $chart.SetSourceData((Resolve-ExcelRange ([string]$actionParams.sourceRange))) }
    if ($actionParams.chartType) { $chart.ChartType = Convert-ExcelChartType ([string]$actionParams.chartType) }
    if ($null -ne $actionParams.name) { $chartObject.Name = [string]$actionParams.name }
    if ($null -ne $actionParams.showTitle) { $chart.HasTitle = [bool]$actionParams.showTitle }
    if ($null -ne $actionParams.title) { $chart.HasTitle = $true; $chart.ChartTitle.Text = [string]$actionParams.title }
    if ($actionParams.style) { try { $chart.ChartStyle = [int]$actionParams.style } catch {} }
    if ($null -ne $actionParams.showLegend) { $chart.HasLegend = [bool]$actionParams.showLegend }
    if ($actionParams.legendPosition -and $chart.HasLegend) {
      $chart.Legend.Position = switch ([string]$actionParams.legendPosition) { 'top' { -4160 } 'bottom' { -4107 } 'left' { -4131 } 'corner' { 2 } default { -4152 } }
    }
    if ($null -ne $actionParams.left) { $chartObject.Left = [double]$actionParams.left }
    if ($null -ne $actionParams.top) { $chartObject.Top = [double]$actionParams.top }
    if ($null -ne $actionParams.width) { $chartObject.Width = [double]$actionParams.width }
    if ($null -ne $actionParams.height) { $chartObject.Height = [double]$actionParams.height }
    if ($actionParams.chartArea) {
      if ($actionParams.chartArea.fillColor) { $chart.ChartArea.Format.Fill.Solid(); $chart.ChartArea.Format.Fill.ForeColor.RGB = Convert-ExcelHexColor $actionParams.chartArea.fillColor 16777215 }
      if ($null -ne $actionParams.chartArea.fillTransparency) { $chart.ChartArea.Format.Fill.Transparency = [double]$actionParams.chartArea.fillTransparency }
      if ($actionParams.chartArea.borderColor) { $chart.ChartArea.Format.Line.Visible = -1; $chart.ChartArea.Format.Line.ForeColor.RGB = Convert-ExcelHexColor $actionParams.chartArea.borderColor 12632256 }
    }
    if ($actionParams.plotArea) {
      if ($actionParams.plotArea.fillColor) { $chart.PlotArea.Format.Fill.Solid(); $chart.PlotArea.Format.Fill.ForeColor.RGB = Convert-ExcelHexColor $actionParams.plotArea.fillColor 16777215 }
      if ($null -ne $actionParams.plotArea.fillTransparency) { $chart.PlotArea.Format.Fill.Transparency = [double]$actionParams.plotArea.fillTransparency }
    }
    if ($actionParams.replaceSeries -eq $true) { while ($chart.SeriesCollection().Count -gt 0) { $chart.SeriesCollection(1).Delete() } }
    foreach ($seriesConfig in @($actionParams.series)) { if ($null -ne $seriesConfig) { Set-ChartSeries $chart $seriesConfig } }
    foreach ($axisConfig in @($actionParams.axes)) { if ($null -ne $axisConfig) { Set-ChartAxis $chart $axisConfig } }
    if ($actionParams.dataLabels) { foreach ($series in $chart.SeriesCollection()) { Set-SeriesDataLabels $series $actionParams.dataLabels } }
    if ($actionParams.exportPath) { [void]$chart.Export([string]$actionParams.exportPath, 'PNG') }
    $snapshot = Get-ChartSnapshot $chartObject ([string]$sheet.Name)
    $operationData.chart = $snapshot
    $changes += [pscustomobject]@{ kind = 'chart-style'; target = [string]$chartObject.Name; detail = '已完成图表深度编辑' }
  }
`;
}
