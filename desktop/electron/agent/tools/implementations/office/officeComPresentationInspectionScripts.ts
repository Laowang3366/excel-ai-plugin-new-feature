const PRESENTATION_INSPECTION_OPERATIONS = new Set([
  "inspectPresentationTheme",
  "inspectSlideElements",
]);

export const PRESENTATION_INSPECTION_HELPERS = String.raw`
  function Get-PresentationShapeTypeName([int]$type) {
    switch ($type) {
      1 { 'autoShape' }
      3 { 'chart' }
      6 { 'group' }
      7 { 'embeddedObject' }
      13 { 'picture' }
      14 { 'placeholder' }
      17 { 'textBox' }
      19 { 'table' }
      default { 'type-' + $type }
    }
  }
  function Get-PresentationShapeText($shape) {
    try {
      if ($shape.HasTextFrame -and $shape.TextFrame.HasText) { return [string]$shape.TextFrame.TextRange.Text }
    } catch {}
    return ''
  }
  function Get-PresentationShapeSnapshot($shape, [double]$slideWidth, [double]$slideHeight) {
    $left = [double]$shape.Left
    $top = [double]$shape.Top
    $width = [double]$shape.Width
    $height = [double]$shape.Height
    $text = Get-PresentationShapeText $shape
    $boundWidth = 0; $boundHeight = 0; $overflow = $false
    if ($text) {
      try {
        $boundWidth = [double]$shape.TextFrame2.TextRange.BoundWidth
        $boundHeight = [double]$shape.TextFrame2.TextRange.BoundHeight
        $availableWidth = [Math]::Max(1, $width - [double]$shape.TextFrame2.MarginLeft - [double]$shape.TextFrame2.MarginRight)
        $availableHeight = [Math]::Max(1, $height - [double]$shape.TextFrame2.MarginTop - [double]$shape.TextFrame2.MarginBottom)
        $overflow = $boundWidth -gt ($availableWidth + 1) -or $boundHeight -gt ($availableHeight + 1)
      } catch {}
    }
    $table = $null
    try {
      if ($shape.HasTable) {
        $cells = @()
        for ($rowIndex = 1; $rowIndex -le $shape.Table.Rows.Count; $rowIndex++) {
          for ($columnIndex = 1; $columnIndex -le $shape.Table.Columns.Count; $columnIndex++) {
            if ($cells.Count -ge 50) { break }
            $cells += [pscustomobject]@{ row = $rowIndex; column = $columnIndex; text = [string]$shape.Table.Cell($rowIndex, $columnIndex).Shape.TextFrame.TextRange.Text }
          }
        }
        $table = [pscustomobject]@{ rows = [int]$shape.Table.Rows.Count; columns = [int]$shape.Table.Columns.Count; cells = $cells }
      }
    } catch {}
    $chart = $null
    try {
      if ($shape.HasChart) {
        $chart = [pscustomobject]@{
          chartType = [int]$shape.Chart.ChartType
          hasTitle = [bool]$shape.Chart.HasTitle
          title = $(try { [string]$shape.Chart.ChartTitle.Text } catch { '' })
          seriesCount = $(try { [int]$shape.Chart.SeriesCollection().Count } catch { 0 })
        }
      }
    } catch {}
    $picture = $null
    if ([int]$shape.Type -eq 13) {
      $picture = [pscustomobject]@{
        cropLeft = $(try { [double]$shape.PictureFormat.CropLeft } catch { 0 })
        cropRight = $(try { [double]$shape.PictureFormat.CropRight } catch { 0 })
        cropTop = $(try { [double]$shape.PictureFormat.CropTop } catch { 0 })
        cropBottom = $(try { [double]$shape.PictureFormat.CropBottom } catch { 0 })
        lockAspectRatio = $(try { [int]$shape.LockAspectRatio } catch { 0 })
      }
    }
    return [pscustomobject]@{
      id = [int]$shape.Id
      name = [string]$shape.Name
      type = [int]$shape.Type
      typeName = Get-PresentationShapeTypeName ([int]$shape.Type)
      zOrder = $(try { [int]$shape.ZOrderPosition } catch { 0 })
      left = $left
      top = $top
      width = $width
      height = $height
      rotation = $(try { [double]$shape.Rotation } catch { 0 })
      text = $text
      textBounds = [pscustomobject]@{ width = $boundWidth; height = $boundHeight }
      textOverflow = $overflow
      outOfBounds = $left -lt 0 -or $top -lt 0 -or ($left + $width) -gt ($slideWidth + 1) -or ($top + $height) -gt ($slideHeight + 1)
      table = $table
      chart = $chart
      picture = $picture
    }
  }
  function Get-SlideElementSnapshot($targetSlide) {
    $slideWidth = [double]$pres.PageSetup.SlideWidth
    $slideHeight = [double]$pres.PageSetup.SlideHeight
    $shapes = @()
    foreach ($shape in $targetSlide.Shapes) { $shapes += Get-PresentationShapeSnapshot $shape $slideWidth $slideHeight }
    $overlaps = @()
    for ($leftIndex = 0; $leftIndex -lt $shapes.Count; $leftIndex++) {
      for ($rightIndex = $leftIndex + 1; $rightIndex -lt $shapes.Count; $rightIndex++) {
        $a = $shapes[$leftIndex]; $b = $shapes[$rightIndex]
        $overlapWidth = [Math]::Min($a.left + $a.width, $b.left + $b.width) - [Math]::Max($a.left, $b.left)
        $overlapHeight = [Math]::Min($a.top + $a.height, $b.top + $b.height) - [Math]::Max($a.top, $b.top)
        if ($overlapWidth -gt 1 -and $overlapHeight -gt 1) {
          $overlaps += [pscustomobject]@{ first = $a.name; second = $b.name; width = $overlapWidth; height = $overlapHeight; area = $overlapWidth * $overlapHeight }
        }
      }
    }
    return [pscustomobject]@{
      index = [int]$targetSlide.SlideIndex
      name = [string]$targetSlide.Name
      layout = $(try { [string]$targetSlide.CustomLayout.Name } catch { '' })
      shapes = $shapes
      overlaps = $overlaps
      overflowCount = @($shapes | Where-Object { $_.textOverflow }).Count
      outOfBoundsCount = @($shapes | Where-Object { $_.outOfBounds }).Count
    }
  }
  function Get-PresentationThemeSnapshot() {
    $designs = @()
    foreach ($design in $pres.Designs) {
      $layouts = @()
      foreach ($layout in $design.SlideMaster.CustomLayouts) {
        $layouts += [pscustomobject]@{ index = [int]$layout.Index; name = [string]$layout.Name; shapeCount = [int]$layout.Shapes.Count }
      }
      $colors = @()
      foreach ($colorIndex in 1..12) {
        try { $colors += [pscustomobject]@{ index = $colorIndex; rgb = [int]$design.SlideMaster.Theme.ThemeColorScheme.Colors($colorIndex).RGB } } catch {}
      }
      $designs += [pscustomobject]@{
        index = [int]$design.Index
        name = [string]$design.Name
        masterName = [string]$design.SlideMaster.Name
        masterShapeCount = [int]$design.SlideMaster.Shapes.Count
        layouts = $layouts
        colors = $colors
      }
    }
    return [pscustomobject]@{
      slideCount = [int]$pres.Slides.Count
      width = [double]$pres.PageSetup.SlideWidth
      height = [double]$pres.PageSetup.SlideHeight
      designs = $designs
      footer = [pscustomobject]@{
        visible = $(try { [bool]$pres.SlideMaster.HeadersFooters.Footer.Visible } catch { $false })
        text = $(try { [string]$pres.SlideMaster.HeadersFooters.Footer.Text } catch { '' })
        slideNumberVisible = $(try { [bool]$pres.SlideMaster.HeadersFooters.SlideNumber.Visible } catch { $false })
      }
    }
  }
`;

export function buildPresentationInspectionOperationScript(operation: string): string | undefined {
  if (!PRESENTATION_INSPECTION_OPERATIONS.has(operation)) return undefined;
  if (operation === "inspectPresentationTheme") {
    return String.raw`
${PRESENTATION_INSPECTION_HELPERS}
  $operationData.theme = Get-PresentationThemeSnapshot
  $changes = @()
`;
  }
  return String.raw`
${PRESENTATION_INSPECTION_HELPERS}
  $targetSlides = if ($actionParams.allSlides -eq $true) { @($pres.Slides) } else { @($slide) }
  $snapshots = @()
  foreach ($targetSlide in $targetSlides) { $snapshots += Get-SlideElementSnapshot $targetSlide }
  $operationData.slides = $snapshots
  $operationData.summary = [pscustomobject]@{
    slideCount = $snapshots.Count
    shapeCount = @($snapshots | ForEach-Object { $_.shapes } | Where-Object { $null -ne $_ }).Count
    overflowCount = @($snapshots | ForEach-Object { $_.shapes } | Where-Object { $_.textOverflow }).Count
    outOfBoundsCount = @($snapshots | ForEach-Object { $_.shapes } | Where-Object { $_.outOfBounds }).Count
    overlapCount = @($snapshots | ForEach-Object { $_.overlaps } | Where-Object { $null -ne $_ }).Count
  }
  $changes = @()
`;
}
