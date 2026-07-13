import { PRESENTATION_INSPECTION_HELPERS } from "./officeComPresentationInspectionScripts";

const PRESENTATION_BRANDING_OPERATIONS = new Set([
  "applyMasterBranding",
  "layoutElements",
]);

export function buildPresentationBrandingOperationScript(operation: string): string | undefined {
  if (!PRESENTATION_BRANDING_OPERATIONS.has(operation)) return undefined;
  return operation === "applyMasterBranding" ? brandingScript() : layoutScript();
}

function brandingScript(): string {
  return String.raw`
${PRESENTATION_INSPECTION_HELPERS}
  function Convert-PresentationColor($value, [int]$fallback) {
    if ($null -eq $value) { return $fallback }
    if ($value -is [ValueType]) { return [int]$value }
    $hex = ([string]$value).Trim().TrimStart('#')
    if ($hex -match '^[0-9A-Fa-f]{6}$') {
      $red = [Convert]::ToInt32($hex.Substring(0, 2), 16)
      $green = [Convert]::ToInt32($hex.Substring(2, 2), 16)
      $blue = [Convert]::ToInt32($hex.Substring(4, 2), 16)
      return $red + ($green * 256) + ($blue * 65536)
    }
    return $fallback
  }
  function Convert-PresentationThemeHex($value) {
    if ($null -eq $value) { return $null }
    $text = ([string]$value).Trim().TrimStart('#')
    if ($text -match '^[0-9A-Fa-f]{6}$') { return $text.ToUpperInvariant() }
    if ($value -is [ValueType]) {
      $ole = [int]$value
      return '{0:X2}{1:X2}{2:X2}' -f ($ole -band 255), (($ole -shr 8) -band 255), (($ole -shr 16) -band 255)
    }
    return $null
  }
  function Update-PresentationThemePackage([string]$path, $rules) {
    if (-not [IO.File]::Exists($path)) { throw '主题色回退更新找不到演示文件: ' + $path }
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = $null
    $lastError = $null
    foreach ($attempt in 1..4) {
      try { $archive = [IO.Compression.ZipFile]::Open($path, [IO.Compression.ZipArchiveMode]::Update); break }
      catch { $lastError = $_; Start-Sleep -Milliseconds (250 * $attempt) }
    }
    if ($null -eq $archive) { throw $lastError }
    try {
      $slots = @{ 1 = 'dk1'; 2 = 'lt1'; 3 = 'dk2'; 4 = 'lt2'; 5 = 'accent1'; 6 = 'accent2'; 7 = 'accent3'; 8 = 'accent4'; 9 = 'accent5'; 10 = 'accent6'; 11 = 'hlink'; 12 = 'folHlink' }
      $entryNames = @($archive.Entries | Where-Object { $_.FullName -match '^ppt/theme/theme[^/]*[.]xml$' } | ForEach-Object { $_.FullName })
      if ($entryNames.Count -eq 0) { throw '演示文稿中没有找到主题 XML' }
      foreach ($entryName in $entryNames) {
        $entry = $archive.GetEntry($entryName)
        $reader = [IO.StreamReader]::new($entry.Open(), [Text.Encoding]::UTF8, $true)
        try { $xmlText = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $xml = [Xml.XmlDocument]::new(); $xml.PreserveWhitespace = $true; $xml.LoadXml($xmlText)
        $ns = [Xml.XmlNamespaceManager]::new($xml.NameTable); $ns.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
        $scheme = $xml.SelectSingleNode('/a:theme/a:themeElements/a:clrScheme', $ns)
        if ($null -eq $scheme) { continue }
        foreach ($rule in @($rules)) {
          $index = [int]$rule.index; $hex = Convert-PresentationThemeHex $rule.value
          if (-not $slots.ContainsKey($index) -or -not $hex) { continue }
          $slot = $scheme.SelectSingleNode('a:' + $slots[$index], $ns)
          if ($null -eq $slot) { continue }
          while ($slot.HasChildNodes) { [void]$slot.RemoveChild($slot.FirstChild) }
          $color = $xml.CreateElement('a', 'srgbClr', 'http://schemas.openxmlformats.org/drawingml/2006/main')
          $color.SetAttribute('val', $hex); [void]$slot.AppendChild($color)
        }
        $entry.Delete(); $replacement = $archive.CreateEntry($entryName, [IO.Compression.CompressionLevel]::Optimal)
        $settings = [Xml.XmlWriterSettings]::new(); $settings.Encoding = [Text.UTF8Encoding]::new($false); $settings.Indent = $false
        $stream = $replacement.Open(); $writer = [Xml.XmlWriter]::Create($stream, $settings)
        try { $xml.Save($writer) } finally { $writer.Dispose(); $stream.Dispose() }
      }
    } finally { $archive.Dispose() }
  }
  function Set-PresentationShapeBrand($shape) {
    try {
      if ($shape.Type -eq 6) {
        foreach ($item in $shape.GroupItems) { Set-PresentationShapeBrand $item }
      }
    } catch {}
    try {
      if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
        $textRange = $shape.TextFrame.TextRange
        if ($actionParams.fontName) { $textRange.Font.Name = [string]$actionParams.fontName }
        if ($actionParams.fontMap) {
          foreach ($mapping in $actionParams.fontMap.PSObject.Properties) {
            try { if ([string]$textRange.Font.Name -eq [string]$mapping.Name) { $textRange.Font.Name = [string]$mapping.Value } } catch {}
          }
        }
        if ($actionParams.applyAccentToText -eq $true) { $textRange.Font.Color.RGB = $accentColor }
      }
    } catch {}
    try {
      if ($shape.HasTable -and $actionParams.fontName) {
        foreach ($row in $shape.Table.Rows) { foreach ($cell in $row.Cells) { $cell.Shape.TextFrame.TextRange.Font.Name = [string]$actionParams.fontName } }
      }
    } catch {}
  }
  function Set-PresentationLogo($master) {
    if (-not $actionParams.logoPath) { return }
    if (-not [IO.File]::Exists([string]$actionParams.logoPath)) { throw '品牌 Logo 不存在: ' + [string]$actionParams.logoPath }
    for ($shapeIndex = $master.Shapes.Count; $shapeIndex -ge 1; $shapeIndex--) {
      $candidate = $master.Shapes.Item($shapeIndex)
      try { if ([string]$candidate.Tags.Item('WENGGE_BRAND_LOGO') -eq '1') { $candidate.Delete() } } catch {}
    }
    $logoWidth = if ($actionParams.logoWidth) { [double]$actionParams.logoWidth } else { 100 }
    $logoHeight = if ($actionParams.logoHeight) { [double]$actionParams.logoHeight } else { -1 }
    $logoLeft = if ($null -ne $actionParams.logoLeft) { [double]$actionParams.logoLeft } else { [double]$pres.PageSetup.SlideWidth - $logoWidth - 24 }
    $logoTop = if ($null -ne $actionParams.logoTop) { [double]$actionParams.logoTop } else { 18 }
    $logo = $master.Shapes.AddPicture([string]$actionParams.logoPath, $false, $true, $logoLeft, $logoTop, $logoWidth, $logoHeight)
    $logo.Name = 'Wengge Brand Logo'
    $logo.Tags.Add('WENGGE_BRAND_LOGO', '1')
    try { $logo.LockAspectRatio = -1 } catch {}
  }

  if ($actionParams.templatePath) {
    if (-not [IO.File]::Exists([string]$actionParams.templatePath)) { throw 'PPT 模板不存在: ' + [string]$actionParams.templatePath }
    $pres.ApplyTemplate([string]$actionParams.templatePath)
    $operationData.appliedTemplate = [string]$actionParams.templatePath
  }
  $updatedMasters = 0; $updatedLayouts = 0; $updatedSlides = 0
  foreach ($design in $pres.Designs) {
    $master = $design.SlideMaster
    if ($actionParams.backgroundColor) {
      try { $master.Background.Fill.Solid(); $master.Background.Fill.ForeColor.RGB = $backgroundColor } catch {}
    }
    if ($actionParams.footerText) {
      try { $master.HeadersFooters.Footer.Visible = $true; $master.HeadersFooters.Footer.Text = [string]$actionParams.footerText } catch {}
    }
    if ($actionParams.showSlideNumber -ne $false) { try { $master.HeadersFooters.SlideNumber.Visible = $true } catch {} }
    foreach ($shape in $master.Shapes) { Set-PresentationShapeBrand $shape }
    foreach ($layout in $master.CustomLayouts) {
      foreach ($shape in $layout.Shapes) { Set-PresentationShapeBrand $shape }
      $updatedLayouts++
    }
    if ($actionParams.themeColors) {
      foreach ($colorRule in @($actionParams.themeColors)) {
        if ($colorRule.index -and $colorRule.value) {
          $targetThemeColor = Convert-PresentationColor $colorRule.value $accentColor
          try {
            $themeColor = $master.Theme.ThemeColorScheme.Colors([int]$colorRule.index)
            $themeColor.RGB = $targetThemeColor
            if ([int]$themeColor.RGB -ne $targetThemeColor) { $pendingThemePackageUpdate = $true }
          } catch { $pendingThemePackageUpdate = $true }
        }
      }
    }
    Set-PresentationLogo $master
    $updatedMasters++
  }
  foreach ($targetSlide in $pres.Slides) {
    foreach ($shape in $targetSlide.Shapes) { Set-PresentationShapeBrand $shape }
    if ($actionParams.footerText) { try { $targetSlide.HeadersFooters.Footer.Visible = $true } catch {} }
    if ($actionParams.showSlideNumber -ne $false) { try { $targetSlide.HeadersFooters.SlideNumber.Visible = $true } catch {} }
    if ($actionParams.layoutMap) {
      $layoutName = $null
      foreach ($mapping in @($actionParams.layoutMap)) {
        if (($mapping.slideIndex -and [int]$mapping.slideIndex -eq [int]$targetSlide.SlideIndex) -or ($mapping.slideName -and [string]$mapping.slideName -eq [string]$targetSlide.Name)) { $layoutName = [string]$mapping.layoutName; break }
      }
      if ($layoutName) {
        foreach ($layout in $targetSlide.Design.SlideMaster.CustomLayouts) { if ([string]$layout.Name -eq $layoutName) { $targetSlide.CustomLayout = $layout; break } }
      }
    }
    $updatedSlides++
  }
  $operationData.updated = [pscustomobject]@{ masters = $updatedMasters; layouts = $updatedLayouts; slides = $updatedSlides }
  $operationData.themePackageFallback = [bool]$pendingThemePackageUpdate
  $operationData.theme = Get-PresentationThemeSnapshot
  $changes += [pscustomobject]@{ kind = 'presentation-brand'; target = 'presentation'; detail = '已统一母版、版式、字体、Logo、页脚和品牌配色' }
`;
}

function layoutScript(): string {
  return String.raw`
${PRESENTATION_INSPECTION_HELPERS}
  function Resolve-LayoutShapes($targetSlide) {
    $selected = @()
    foreach ($shape in $targetSlide.Shapes) {
      if ($actionParams.shapeNames -and @($actionParams.shapeNames) -notcontains [string]$shape.Name) { continue }
      if ($actionParams.excludePlaceholders -ne $false) {
        try { if ($shape.Type -eq 14 -and $shape.PlaceholderFormat.Type -eq 1) { continue } } catch {}
      }
      $selected += $shape
    }
    return $selected
  }
  function Apply-PreciseShapeEdits($targetSlide) {
    $count = 0
    foreach ($edit in @($actionParams.edits)) {
      $shape = $null
      if ($edit.shapeName) { try { $shape = $targetSlide.Shapes.Item([string]$edit.shapeName) } catch {} }
      elseif ($edit.shapeIndex) { try { $shape = $targetSlide.Shapes.Item([int]$edit.shapeIndex) } catch {} }
      if ($null -eq $shape) { continue }
      $shapeEdited = $false
      if ($edit.preserveAspectRatio -eq $true) { try { $shape.LockAspectRatio = -1; $shapeEdited = $true } catch { $script:layoutEditFailures += [pscustomobject]@{ shape = [string]$shape.Name; property = 'lockAspectRatio'; error = $_.Exception.Message } } }
      if ($null -ne $edit.left) { try { $shape.Left = [single]$edit.left; $shapeEdited = $true } catch { $script:layoutEditFailures += [pscustomobject]@{ shape = [string]$shape.Name; property = 'left'; error = $_.Exception.Message } } }
      if ($null -ne $edit.top) { try { $shape.Top = [single]$edit.top; $shapeEdited = $true } catch { $script:layoutEditFailures += [pscustomobject]@{ shape = [string]$shape.Name; property = 'top'; error = $_.Exception.Message } } }
      if ($null -ne $edit.width) { try { $shape.Width = [single]$edit.width; $shapeEdited = $true } catch { $script:layoutEditFailures += [pscustomobject]@{ shape = [string]$shape.Name; property = 'width'; error = $_.Exception.Message } } }
      if ($null -ne $edit.height) { try { $shape.Height = [single]$edit.height; $shapeEdited = $true } catch { $script:layoutEditFailures += [pscustomobject]@{ shape = [string]$shape.Name; property = 'height'; error = $_.Exception.Message } } }
      if ($null -ne $edit.rotation) { try { $shape.Rotation = [single]$edit.rotation; $shapeEdited = $true } catch { $script:layoutEditFailures += [pscustomobject]@{ shape = [string]$shape.Name; property = 'rotation'; error = $_.Exception.Message } } }
      if ($null -ne $edit.text) { try { $shape.TextFrame.TextRange.Text = [string]$edit.text; $shapeEdited = $true } catch {} }
      if ($edit.fontName) { try { $shape.TextFrame.TextRange.Font.Name = [string]$edit.fontName; $shapeEdited = $true } catch {} }
      if ($edit.fontSize) { try { $shape.TextFrame.TextRange.Font.Size = [double]$edit.fontSize; $shapeEdited = $true } catch {} }
      if ($edit.tableCells) {
        foreach ($cellEdit in @($edit.tableCells)) {
          try {
            if (-not $shape.HasTable) { continue }
            $cell = $shape.Table.Cell([int]$cellEdit.row, [int]$cellEdit.column).Shape
            if ($null -ne $cellEdit.text) { $cell.TextFrame.TextRange.Text = [string]$cellEdit.text }
            if ($cellEdit.fontName) { $cell.TextFrame.TextRange.Font.Name = [string]$cellEdit.fontName }
            if ($cellEdit.fontSize) { $cell.TextFrame.TextRange.Font.Size = [double]$cellEdit.fontSize }
            if ($null -ne $cellEdit.fillColor) { $cell.Fill.Solid(); $cell.Fill.ForeColor.RGB = [int]$cellEdit.fillColor }
          } catch {}
        }
      }
      if ($edit.chart) {
        try {
          if ($shape.HasChart) {
            if ($edit.chart.chartType) { $shape.Chart.ChartType = [int]$edit.chart.chartType }
            if ($null -ne $edit.chart.title) { $shape.Chart.HasTitle = $true; $shape.Chart.ChartTitle.Text = [string]$edit.chart.title }
            if ($null -ne $edit.chart.hasLegend) { $shape.Chart.HasLegend = [bool]$edit.chart.hasLegend }
          }
        } catch {}
      }
      if ($edit.crop) {
        try {
          if ($null -ne $edit.crop.left) { $shape.PictureFormat.CropLeft = [double]$edit.crop.left }
          if ($null -ne $edit.crop.right) { $shape.PictureFormat.CropRight = [double]$edit.crop.right }
          if ($null -ne $edit.crop.top) { $shape.PictureFormat.CropTop = [double]$edit.crop.top }
          if ($null -ne $edit.crop.bottom) { $shape.PictureFormat.CropBottom = [double]$edit.crop.bottom }
        } catch {}
      }
      if ($shapeEdited) { $count++ }
    }
    return $count
  }
  function Align-LayoutShapes($items, [string]$alignment) {
    if ($items.Count -eq 0) { return }
    $minLeft = ($items | Measure-Object Left -Minimum).Minimum
    $maxRight = ($items | ForEach-Object { $_.Left + $_.Width } | Measure-Object -Maximum).Maximum
    $minTop = ($items | Measure-Object Top -Minimum).Minimum
    $maxBottom = ($items | ForEach-Object { $_.Top + $_.Height } | Measure-Object -Maximum).Maximum
    foreach ($item in $items) {
      switch ($alignment) {
        'left' { $item.Left = [single]$minLeft }
        'center' { $item.Left = [single]((($minLeft + $maxRight) - $item.Width) / 2) }
        'right' { $item.Left = [single]($maxRight - $item.Width) }
        'top' { $item.Top = [single]$minTop }
        'middle' { $item.Top = [single]((($minTop + $maxBottom) - $item.Height) / 2) }
        'bottom' { $item.Top = [single]($maxBottom - $item.Height) }
      }
    }
  }
  function Distribute-LayoutShapes($items, [string]$direction) {
    if ($items.Count -lt 3) { return }
    if ($direction -eq 'vertical') {
      $ordered = @($items | Sort-Object Top)
      $totalSize = ($ordered | Measure-Object Height -Sum).Sum
      $gap = (($ordered[-1].Top + $ordered[-1].Height) - $ordered[0].Top - $totalSize) / ($ordered.Count - 1)
      $cursor = [double]$ordered[0].Top
      foreach ($item in $ordered) { $item.Top = [single]$cursor; $cursor += [double]$item.Height + $gap }
    } else {
      $ordered = @($items | Sort-Object Left)
      $totalSize = ($ordered | Measure-Object Width -Sum).Sum
      $gap = (($ordered[-1].Left + $ordered[-1].Width) - $ordered[0].Left - $totalSize) / ($ordered.Count - 1)
      $cursor = [double]$ordered[0].Left
      foreach ($item in $ordered) { $item.Left = [single]$cursor; $cursor += [double]$item.Width + $gap }
    }
  }

  $targetSlides = if ($actionParams.allSlides -eq $true) { @($pres.Slides) } else { @($slide) }
  $mode = if ($actionParams.mode) { [string]$actionParams.mode } else { 'grid' }
  $edited = 0
  $script:layoutEditFailures = @()
  foreach ($targetSlide in $targetSlides) {
    $edited += Apply-PreciseShapeEdits $targetSlide
    $items = @(Resolve-LayoutShapes $targetSlide)
    if ($mode -eq 'grid' -or $mode -eq 'auto') {
      $columns = if ($actionParams.columns) { [Math]::Max(1, [int]$actionParams.columns) } else { 2 }
      $margin = if ($actionParams.margin) { [double]$actionParams.margin } else { 40 }
      $gap = if ($actionParams.gap) { [double]$actionParams.gap } else { 16 }
      $cellWidth = ($pres.PageSetup.SlideWidth - (2 * $margin) - (($columns - 1) * $gap)) / $columns
      $rowHeight = if ($actionParams.rowHeight) { [double]$actionParams.rowHeight } else { 140 }
      for ($itemIndex = 0; $itemIndex -lt $items.Count; $itemIndex++) {
        $column = $itemIndex % $columns; $row = [Math]::Floor($itemIndex / $columns)
        $items[$itemIndex].Left = [single]($margin + ($column * ($cellWidth + $gap)))
        $items[$itemIndex].Top = [single]($margin + ($row * ($rowHeight + $gap)))
        if ($actionParams.resize -ne $false) {
          if ($actionParams.preserveAspectRatio -eq $true) { try { $items[$itemIndex].LockAspectRatio = -1 } catch {} }
          $items[$itemIndex].Width = [single]$cellWidth
          if ($items[$itemIndex].Height -gt $rowHeight) { $items[$itemIndex].Height = [single]$rowHeight }
        }
      }
    }
    if ($actionParams.align) { Align-LayoutShapes $items ([string]$actionParams.align) }
    if ($actionParams.distribute) { Distribute-LayoutShapes $items ([string]$actionParams.distribute) }
    if ($actionParams.fitToSlide -eq $true) {
      foreach ($item in $items) {
        if ($item.Left -lt 0) { $item.Left = [single]0 }; if ($item.Top -lt 0) { $item.Top = [single]0 }
        if (($item.Left + $item.Width) -gt $pres.PageSetup.SlideWidth) { $item.Left = [single]([Math]::Max(0, $pres.PageSetup.SlideWidth - $item.Width)) }
        if (($item.Top + $item.Height) -gt $pres.PageSetup.SlideHeight) { $item.Top = [single]([Math]::Max(0, $pres.PageSetup.SlideHeight - $item.Height)) }
      }
    }
  }
  $snapshots = @(); foreach ($targetSlide in $targetSlides) { $snapshots += Get-SlideElementSnapshot $targetSlide }
  $operationData.editedShapes = $edited
  $operationData.editFailures = $script:layoutEditFailures
  $operationData.slides = $snapshots
  $operationData.summary = [pscustomobject]@{
    overflowCount = @($snapshots | ForEach-Object { $_.shapes } | Where-Object { $_.textOverflow }).Count
    outOfBoundsCount = @($snapshots | ForEach-Object { $_.shapes } | Where-Object { $_.outOfBounds }).Count
    overlapCount = @($snapshots | ForEach-Object { $_.overlaps } | Where-Object { $null -ne $_ }).Count
  }
  $changes += [pscustomobject]@{ kind = 'presentation-layout'; target = 'slides'; detail = '已精确编辑、对齐、等距分布、裁剪并检查元素布局' }
`;
}
