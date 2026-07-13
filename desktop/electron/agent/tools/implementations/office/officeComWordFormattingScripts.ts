const WORD_FORMATTING_OPERATIONS = new Set([
  "applyHeadingStyles",
  "inspectDocumentFormatting",
  "formatLongDocument",
]);

export function buildWordFormattingOperationScript(operation: string): string | undefined {
  if (!WORD_FORMATTING_OPERATIONS.has(operation)) return undefined;
  return String.raw`
  function Get-StyleSnapshot($style, [string]$name) {
    return [pscustomobject]@{
      name = $name
      fontName = $(try { [string]$style.Font.Name } catch { '' })
      fontSize = $(try { [double]$style.Font.Size } catch { 0 })
      bold = $(try { [bool]$style.Font.Bold } catch { $false })
      italic = $(try { [bool]$style.Font.Italic } catch { $false })
      alignment = $(try { [int]$style.ParagraphFormat.Alignment } catch { 0 })
      outlineLevel = $(try { [int]$style.ParagraphFormat.OutlineLevel } catch { 10 })
      spaceBefore = $(try { [double]$style.ParagraphFormat.SpaceBefore } catch { 0 })
      spaceAfter = $(try { [double]$style.ParagraphFormat.SpaceAfter } catch { 0 })
      lineSpacingRule = $(try { [int]$style.ParagraphFormat.LineSpacingRule } catch { 0 })
    }
  }
  function Set-WordStyle($style, $rule, [string]$defaultFont, [double]$defaultSize) {
    $style.Font.Name = if ($rule.fontName) { [string]$rule.fontName } else { $defaultFont }
    $style.Font.Size = if ($rule.fontSize) { [double]$rule.fontSize } else { $defaultSize }
    if ($null -ne $rule.bold) { $style.Font.Bold = [bool]$rule.bold }
    if ($null -ne $rule.italic) { $style.Font.Italic = [bool]$rule.italic }
    if ($rule.fontColor) { try { $style.Font.Color = [int]$rule.fontColor } catch {} }
    if ($null -ne $rule.alignment) { $style.ParagraphFormat.Alignment = [int]$rule.alignment }
    if ($null -ne $rule.spaceBefore) { $style.ParagraphFormat.SpaceBefore = [double]$rule.spaceBefore }
    if ($null -ne $rule.spaceAfter) { $style.ParagraphFormat.SpaceAfter = [double]$rule.spaceAfter }
    if ($null -ne $rule.firstLineIndent) { $style.ParagraphFormat.FirstLineIndent = $app.CentimetersToPoints([double]$rule.firstLineIndent) }
    if ($null -ne $rule.leftIndent) { $style.ParagraphFormat.LeftIndent = $app.CentimetersToPoints([double]$rule.leftIndent) }
    if ($null -ne $rule.rightIndent) { $style.ParagraphFormat.RightIndent = $app.CentimetersToPoints([double]$rule.rightIndent) }
    if ($null -ne $rule.lineSpacing) {
      $style.ParagraphFormat.LineSpacingRule = 5
      $style.ParagraphFormat.LineSpacing = $app.LinesToPoints([double]$rule.lineSpacing)
    }
    if ($null -ne $rule.keepWithNext) { $style.ParagraphFormat.KeepWithNext = [bool]$rule.keepWithNext }
    if ($null -ne $rule.keepTogether) { $style.ParagraphFormat.KeepTogether = [bool]$rule.keepTogether }
  }
  function Get-HeadingLevelFromText([string]$text) {
    foreach ($rule in @($actionParams.headingRules)) {
      if ($rule.pattern -and [regex]::IsMatch($text, [string]$rule.pattern)) { return [Math]::Max(1, [Math]::Min(9, [int]$rule.level)) }
    }
    if ($actionParams.autoDetectHeadings -eq $false) { return 0 }
    if ($text -match '^\s*(第[一二三四五六七八九十百]+[章节篇]|[一二三四五六七八九十]+、|\d+[、.．]\s*[^\d])') { return 1 }
    if ($text -match '^\s*([（(][一二三四五六七八九十\d]+[）)]|\d+\.\d+[、.．]?)') { return 2 }
    if ($text -match '^\s*(\d+\.\d+\.\d+|[①②③④⑤⑥⑦⑧⑨⑩])') { return 3 }
    return 0
  }
  function Test-ParagraphInTable($paragraph) {
    try { return [bool]$paragraph.Range.Information(12) } catch { return $false }
  }
  function Set-HeaderFooterText($section, $config) {
    if ($null -eq $config) { return }
    if ($null -ne $config.header) { $section.Headers.Item(1).Range.Text = [string]$config.header }
    if ($null -ne $config.footer) { $section.Footers.Item(1).Range.Text = [string]$config.footer }
    if ($null -ne $config.firstPageHeader) {
      $section.PageSetup.DifferentFirstPageHeaderFooter = $true
      $section.Headers.Item(2).Range.Text = [string]$config.firstPageHeader
    }
    if ($null -ne $config.firstPageFooter) {
      $section.PageSetup.DifferentFirstPageHeaderFooter = $true
      $section.Footers.Item(2).Range.Text = [string]$config.firstPageFooter
    }
    if ($null -ne $config.differentOddEven) { $section.PageSetup.OddAndEvenPagesHeaderFooter = [bool]$config.differentOddEven }
  }
  function Update-WordFields() {
    foreach ($story in $doc.StoryRanges) {
      $current = $story
      while ($null -ne $current) {
        try { if ($current.Fields.Count -gt 0) { $current.Fields.Update() } } catch {}
        try { $current = $current.NextStoryRange } catch { $current = $null }
      }
    }
    foreach ($toc in $doc.TablesOfContents) { try { $toc.Update() } catch {} }
    foreach ($figureTable in $doc.TablesOfFigures) { try { $figureTable.Update() } catch {} }
  }
  function Get-DocumentFormattingSnapshot() {
    $styles = @()
    $styles += Get-StyleSnapshot $doc.Styles.Item(-1) 'Normal'
    foreach ($level in 1..9) {
      try { $styles += Get-StyleSnapshot $doc.Styles.Item(-1 - $level) ('Heading' + $level) } catch {}
    }
    try { $styles += Get-StyleSnapshot $doc.Styles.Item(-35) 'Caption' } catch {}
    $styleCounts = @{}
    $emptyParagraphs = 0
    foreach ($paragraph in $doc.Paragraphs) {
      $text = ([string]$paragraph.Range.Text).Trim()
      if (-not $text) { $emptyParagraphs++; continue }
      $styleName = try { [string]$paragraph.Range.Style.NameLocal } catch { [string]$paragraph.Range.Style }
      if (-not $styleCounts.ContainsKey($styleName)) { $styleCounts[$styleName] = 0 }
      $styleCounts[$styleName]++
    }
    $styleUsage = @()
    foreach ($name in $styleCounts.Keys) { $styleUsage += [pscustomobject]@{ style = $name; paragraphs = $styleCounts[$name] } }
    $sections = @()
    foreach ($section in $doc.Sections) {
      $sections += [pscustomobject]@{
        index = [int]$section.Index
        start = [int]$section.Range.Start
        end = [int]$section.Range.End
        orientation = [int]$section.PageSetup.Orientation
        marginsPoints = [pscustomobject]@{ top = [double]$section.PageSetup.TopMargin; bottom = [double]$section.PageSetup.BottomMargin; left = [double]$section.PageSetup.LeftMargin; right = [double]$section.PageSetup.RightMargin }
        header = $(try { [string]$section.Headers.Item(1).Range.Text.Trim() } catch { '' })
        footer = $(try { [string]$section.Footers.Item(1).Range.Text.Trim() } catch { '' })
        pageNumbers = $(try { [int]$section.Footers.Item(1).PageNumbers.Count } catch { 0 })
      }
    }
    return [pscustomobject]@{
      styles = $styles
      styleUsage = $styleUsage
      sections = $sections
      paragraphCount = [int]$doc.Paragraphs.Count
      emptyParagraphCount = $emptyParagraphs
      tableCount = [int]$doc.Tables.Count
      inlineShapeCount = [int]$doc.InlineShapes.Count
      shapeCount = [int]$doc.Shapes.Count
      tocCount = [int]$doc.TablesOfContents.Count
      fieldCount = [int]$doc.Fields.Count
    }
  }

  if ($_operation -eq 'inspectDocumentFormatting') {
    $snapshot = Get-DocumentFormattingSnapshot
    $operationData.formatting = $snapshot
    $changes = @()
  } elseif ($_operation -eq 'applyHeadingStyles') {
    $styled = 0
    foreach ($paragraph in $doc.Paragraphs) {
      $text = ([string]$paragraph.Range.Text).Trim()
      if (-not $text) { continue }
      $level = if ($_startsWith -and $text.StartsWith($_startsWith)) { $headingLevel } else { Get-HeadingLevelFromText $text }
      if ($level -gt 0) { $paragraph.Range.Style = $doc.Styles.Item(-1 - $level); $styled++ }
    }
    $operationData.styledHeadings = $styled
    $changes += [pscustomobject]@{ kind = 'document-style'; target = 'headings'; detail = '已自动设置 ' + $styled + ' 个标题层级' }
  } else {
    $normalRule = if ($actionParams.normalStyle) { $actionParams.normalStyle } else { [pscustomobject]@{} }
    $fontName = if ($actionParams.fontName) { [string]$actionParams.fontName } else { '微软雅黑' }
    $fontSize = if ($actionParams.fontSize) { [double]$actionParams.fontSize } else { 10.5 }
    Set-WordStyle $doc.Styles.Item(-1) $normalRule $fontName $fontSize
    if ($null -eq $normalRule.alignment) { $doc.Styles.Item(-1).ParagraphFormat.Alignment = 0 }
    if ($null -eq $normalRule.firstLineIndent) { $doc.Styles.Item(-1).ParagraphFormat.FirstLineIndent = $app.CentimetersToPoints(0.74) }
    if ($null -eq $normalRule.lineSpacing) { $doc.Styles.Item(-1).ParagraphFormat.LineSpacingRule = 1 }
    $headingSizes = @(18, 16, 14, 12, 11, 10.5, 10.5, 10.5, 10.5)
    foreach ($level in 1..9) {
      $rule = if ($actionParams.headingStyles -and $actionParams.headingStyles.PSObject.Properties[[string]$level]) { $actionParams.headingStyles.PSObject.Properties[[string]$level].Value } else { [pscustomobject]@{} }
      $style = $doc.Styles.Item(-1 - $level)
      Set-WordStyle $style $rule $(if ($actionParams.headingFontName) { [string]$actionParams.headingFontName } else { $fontName }) $headingSizes[$level - 1]
      $style.Font.Bold = $true
      $style.ParagraphFormat.KeepWithNext = $true
      $style.ParagraphFormat.OutlineLevel = $level
    }
    try {
      $captionStyle = $doc.Styles.Item(-35)
      $captionStyle.Font.Name = $fontName
      $captionStyle.Font.Size = if ($actionParams.captionFontSize) { [double]$actionParams.captionFontSize } else { 9 }
      $captionStyle.ParagraphFormat.Alignment = 1
      $captionStyle.ParagraphFormat.KeepWithNext = $true
    } catch {}
    $headings = 0; $body = 0; $quotes = 0; $captions = 0
    foreach ($paragraph in $doc.Paragraphs) {
      $text = ([string]$paragraph.Range.Text).Trim()
      if (-not $text -or (Test-ParagraphInTable $paragraph)) { continue }
      $level = Get-HeadingLevelFromText $text
      if ($level -gt 0) {
        $paragraph.Range.Style = $doc.Styles.Item(-1 - $level)
        $headings++
        continue
      }
      if ($text -match '^\s*(图|表|公式)\s*[一二三四五六七八九十\d]+[.．、:：-]?') {
        try { $paragraph.Range.Style = $doc.Styles.Item(-35) } catch {}
        $captions++
        continue
      }
      $isQuote = $false
      foreach ($pattern in @($actionParams.quotePatterns)) { if ($pattern -and [regex]::IsMatch($text, [string]$pattern)) { $isQuote = $true; break } }
      if ($isQuote) {
        try { $paragraph.Range.Style = $doc.Styles.Item('Quote') } catch { $paragraph.Range.Italic = $true; $paragraph.Range.ParagraphFormat.LeftIndent = $app.CentimetersToPoints(1) }
        $quotes++
      } else {
        if ($actionParams.clearDirectFormatting -eq $true) { try { $paragraph.Range.Font.Reset(); $paragraph.Range.ParagraphFormat.Reset() } catch {} }
        $paragraph.Range.Style = $doc.Styles.Item(-1)
        $body++
      }
    }
    $marginUnit = if ($actionParams.marginUnit -eq 'points') { 'points' } else { 'centimeters' }
    foreach ($section in $doc.Sections) {
      $margins = $actionParams.margins
      if ($null -ne $margins) {
        if ($null -ne $margins.top) { $section.PageSetup.TopMargin = if ($marginUnit -eq 'points') { [double]$margins.top } else { $app.CentimetersToPoints([double]$margins.top) } }
        if ($null -ne $margins.bottom) { $section.PageSetup.BottomMargin = if ($marginUnit -eq 'points') { [double]$margins.bottom } else { $app.CentimetersToPoints([double]$margins.bottom) } }
        if ($null -ne $margins.left) { $section.PageSetup.LeftMargin = if ($marginUnit -eq 'points') { [double]$margins.left } else { $app.CentimetersToPoints([double]$margins.left) } }
        if ($null -ne $margins.right) { $section.PageSetup.RightMargin = if ($marginUnit -eq 'points') { [double]$margins.right } else { $app.CentimetersToPoints([double]$margins.right) } }
      }
      if ($null -ne $actionParams.orientation) { $section.PageSetup.Orientation = if ([string]$actionParams.orientation -eq 'landscape') { 1 } else { 0 } }
      Set-HeaderFooterText $section $actionParams.headerFooter
      if ($actionParams.pageNumbers -ne $false -and $section.Footers.Item(1).PageNumbers.Count -eq 0) { [void]$section.Footers.Item(1).PageNumbers.Add(1, $true) }
    }
    foreach ($break in @($actionParams.sectionBreaks | Sort-Object { [int]$_.position } -Descending)) {
      $position = [Math]::Max(0, [Math]::Min([int]$break.position, $doc.Content.End - 1))
      $breakType = switch ([string]$break.type) { 'continuous' { 3 } 'evenPage' { 4 } 'oddPage' { 5 } default { 2 } }
      $doc.Range($position, $position).InsertBreak($breakType)
    }
    foreach ($table in $doc.Tables) {
      $table.Borders.Enable = 1
      try { $table.AutoFitBehavior(1) } catch {}
      if ($table.Rows.Count -gt 0) { $table.Rows.Item(1).Range.Bold = $true }
    }
    if ($actionParams.toc -eq 'create' -and $doc.TablesOfContents.Count -eq 0) {
      $tocPosition = if ($null -ne $actionParams.tocPosition) { [int]$actionParams.tocPosition } else { 0 }
      $toc = $doc.TablesOfContents.Add($doc.Range($tocPosition, $tocPosition), $true, 1, 3)
      $toc.Update()
    } elseif ($actionParams.toc -in @('create', 'update')) {
      foreach ($toc in $doc.TablesOfContents) { $toc.Update() }
    }
    Update-WordFields
    $operationData.styled = [pscustomobject]@{ headings = $headings; body = $body; quotes = $quotes; captions = $captions; sections = [int]$doc.Sections.Count }
    $operationData.formatting = Get-DocumentFormattingSnapshot
    $changes += [pscustomobject]@{ kind = 'document-style'; target = 'document'; detail = '已统一标题、正文、引用、题注和页面结构' }
  }
`;
}
