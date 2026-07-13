const WORD_REFERENCE_OPERATIONS = new Set([
  "insertOrUpdateToc",
  "inspectReferences",
  "manageReferences",
]);

export function buildWordReferenceOperationScript(operation: string): string | undefined {
  if (!WORD_REFERENCE_OPERATIONS.has(operation)) return undefined;
  return String.raw`
  function Get-ReferenceRange() {
    if ($actionParams.bookmarkName -and $doc.Bookmarks.Exists([string]$actionParams.bookmarkName)) {
      $resolved = $doc.Bookmarks.Item([string]$actionParams.bookmarkName).Range
      Write-Output -NoEnumerate $resolved
      return
    }
    $start = if ($null -ne $actionParams.start) { [int]$actionParams.start } elseif ($null -ne $actionParams.position) { [int]$actionParams.position } else { [Math]::Max(0, $doc.Content.End - 1) }
    $end = if ($null -ne $actionParams.end) { [int]$actionParams.end } else { $start }
    $start = [Math]::Max(0, [Math]::Min($start, $doc.Content.End - 1))
    $end = [Math]::Max($start, [Math]::Min($end, $doc.Content.End - 1))
    $resolved = $doc.Range($start, $end)
    Write-Output -NoEnumerate $resolved
  }
  function Get-CaptionTargetRange() {
    $kind = if ($actionParams.targetType) { [string]$actionParams.targetType } else { 'range' }
    $index = if ($actionParams.index) { [Math]::Max(1, [int]$actionParams.index) } else { 1 }
    if ($kind -eq 'table') {
      if ($index -gt $doc.Tables.Count) { throw '题注目标表格不存在' }
      $resolved = $doc.Tables.Item($index).Range
    } elseif ($kind -in @('image', 'figure')) {
      if ($index -le $doc.InlineShapes.Count) { $resolved = $doc.InlineShapes.Item($index).Range }
      elseif ($index -le $doc.Shapes.Count) { $resolved = $doc.Shapes.Item($index).Anchor }
      else { throw '题注目标图片不存在' }
    } elseif ($kind -eq 'equation') {
      if ($index -gt $doc.OMaths.Count) { throw '题注目标公式不存在' }
      $resolved = $doc.OMaths.Item($index).Range
    } else {
      $resolved = Get-ReferenceRange
    }
    Write-Output -NoEnumerate $resolved
  }
  function Resolve-ReferenceType($value) {
    switch ([string]$value) {
      'bookmark' { 0 }
      'heading' { 1 }
      'footnote' { 2 }
      'endnote' { 3 }
      'numberedItem' { -1 }
      'table' { if ($actionParams.label) { [string]$actionParams.label } else { '表' } }
      'figure' { if ($actionParams.label) { [string]$actionParams.label } else { '图' } }
      'equation' { if ($actionParams.label) { [string]$actionParams.label } else { '公式' } }
      default { if ($null -ne $value) { $value } else { 0 } }
    }
  }
  function Resolve-ReferenceKind($referenceType, $value) {
    if ($null -ne $value) { return [int]$value }
    switch ([string]$referenceType) {
      'bookmark' { -1 }
      'heading' { -1 }
      'footnote' { 5 }
      'endnote' { 6 }
      default { 3 }
    }
  }
  function Update-AllWordReferences() {
    $referenceFieldTypes = @(3, 5, 9, 10, 11, 12, 13, 37, 72)
    foreach ($story in $doc.StoryRanges) {
      $current = $story
      while ($null -ne $current) {
        try {
          foreach ($field in $current.Fields) {
            if ([int]$field.Type -in $referenceFieldTypes) { [void]$field.Update() }
          }
        } catch {}
        try { $current = $current.NextStoryRange } catch { $current = $null }
      }
    }
    foreach ($toc in $doc.TablesOfContents) { try { $toc.Update() } catch {} }
    foreach ($tof in $doc.TablesOfFigures) { try { $tof.Update() } catch {} }
  }
  function Get-ReferenceSnapshot() {
    $bookmarks = @()
    foreach ($bookmark in $doc.Bookmarks) { $bookmarks += [pscustomobject]@{ name = [string]$bookmark.Name; start = [int]$bookmark.Range.Start; end = [int]$bookmark.Range.End; text = ([string]$bookmark.Range.Text).Trim() } }
    $footnotes = @()
    foreach ($note in $doc.Footnotes) { $footnotes += [pscustomobject]@{ index = [int]$note.Index; text = ([string]$note.Range.Text).Trim(); referenceStart = [int]$note.Reference.Start } }
    $endnotes = @()
    foreach ($note in $doc.Endnotes) { $endnotes += [pscustomobject]@{ index = [int]$note.Index; text = ([string]$note.Range.Text).Trim(); referenceStart = [int]$note.Reference.Start } }
    $fields = @()
    foreach ($field in $doc.Fields) {
      $fields += [pscustomobject]@{ type = [int]$field.Type; code = ([string]$field.Code.Text).Trim(); result = ([string]$field.Result.Text).Trim(); start = [int]$field.Result.Start; end = [int]$field.Result.End }
    }
    return [pscustomobject]@{
      bookmarks = $bookmarks
      footnotes = $footnotes
      endnotes = $endnotes
      fields = $fields
      tocCount = [int]$doc.TablesOfContents.Count
      tableOfFiguresCount = [int]$doc.TablesOfFigures.Count
      captionLabels = @($app.CaptionLabels | ForEach-Object { [string]$_.Name })
    }
  }

  if ($_operation -eq 'inspectReferences') {
    $operationData.references = Get-ReferenceSnapshot
    $changes = @()
  } elseif ($_operation -eq 'insertOrUpdateToc') {
    if ($doc.TablesOfContents.Count -eq 0) {
      $position = if ($null -ne $actionParams.position) { [int]$actionParams.position } else { 0 }
      $toc = $doc.TablesOfContents.Add($doc.Range($position, $position), $true, $(if ($actionParams.lowerHeadingLevel) { [int]$actionParams.lowerHeadingLevel } else { 1 }), $(if ($actionParams.upperHeadingLevel) { [int]$actionParams.upperHeadingLevel } else { 3 }))
      $toc.Update()
    } else { foreach ($toc in $doc.TablesOfContents) { $toc.Update() } }
    Update-AllWordReferences
    $operationData.tocCount = [int]$doc.TablesOfContents.Count
    $changes += [pscustomobject]@{ kind = 'toc'; target = 'document'; detail = '已创建或更新目录及全部引用域' }
  } else {
    $command = if ($actionParams.command) { [string]$actionParams.command } else { 'updateAll' }
    $missing = [Type]::Missing
    switch ($command) {
      'addFootnote' {
        if (-not $actionParams.text) { throw 'addFootnote 需要 params.text' }
        $target = Get-ReferenceRange
        [void]$doc.Footnotes.Add($target, $missing, [string]$actionParams.text)
      }
      'addEndnote' {
        if (-not $actionParams.text) { throw 'addEndnote 需要 params.text' }
        $target = Get-ReferenceRange
        [void]$doc.Endnotes.Add($target, $missing, [string]$actionParams.text)
      }
      'addBookmark' {
        if (-not $actionParams.name) { throw 'addBookmark 需要 params.name' }
        $target = Get-ReferenceRange
        if ($doc.Bookmarks.Exists([string]$actionParams.name)) { $doc.Bookmarks.Item([string]$actionParams.name).Delete() }
        [void]$doc.Bookmarks.Add([string]$actionParams.name, $target)
      }
      'deleteBookmark' {
        if (-not $actionParams.name) { throw 'deleteBookmark 需要 params.name' }
        if ($doc.Bookmarks.Exists([string]$actionParams.name)) { $doc.Bookmarks.Item([string]$actionParams.name).Delete() }
      }
      'addCaption' {
        $label = if ($actionParams.label) { [string]$actionParams.label } else { switch ([string]$actionParams.targetType) { 'table' { '表' } 'equation' { '公式' } default { '图' } } }
        try { $null = $app.CaptionLabels.Item($label) } catch { [void]$app.CaptionLabels.Add($label) }
        $target = Get-CaptionTargetRange
        $position = if ($actionParams.position -eq 'above') { 0 } elseif ($actionParams.position -eq 'below') { 1 } elseif ($actionParams.targetType -eq 'table') { 0 } else { 1 }
        $title = if ($actionParams.title) { ' ' + [string]$actionParams.title } else { '' }
        $target.InsertCaption($label, $title, $missing, $position, $false)
      }
      'addCrossReference' {
        if ($null -eq $actionParams.item) { throw 'addCrossReference 需要 params.item' }
        $target = Get-ReferenceRange
        $referenceType = Resolve-ReferenceType $actionParams.referenceType
        $referenceKind = Resolve-ReferenceKind $actionParams.referenceType $actionParams.referenceKind
        if ([string]$actionParams.referenceType -eq 'bookmark') {
          if (-not $doc.Bookmarks.Exists([string]$actionParams.item)) { throw '交叉引用书签不存在: ' + [string]$actionParams.item }
          $fieldCode = [string]$actionParams.item
          if ($actionParams.insertAsHyperlink -ne $false) { $fieldCode += ' \h' }
          if ($actionParams.includePosition -eq $true) { $fieldCode += ' \p' }
          $field = $doc.Fields.Add($target, 3, $fieldCode, $true)
          [void]$field.Update()
        } else {
          $target.InsertCrossReference($referenceType, $referenceKind, $actionParams.item, $actionParams.insertAsHyperlink -ne $false, $actionParams.includePosition -eq $true, $actionParams.separateNumbers -eq $true)
        }
      }
      'addTableOfFigures' {
        $target = Get-ReferenceRange
        $label = if ($actionParams.label) { [string]$actionParams.label } else { '图' }
        [void]$doc.TablesOfFigures.Add($target, $label, $true)
      }
      'updateFields' { Update-AllWordReferences }
      'updateAll' { Update-AllWordReferences }
      default { throw '不支持的引用命令: ' + $command }
    }
    if ($command -in @('addCaption', 'addCrossReference', 'addTableOfFigures')) { Update-AllWordReferences }
    $operationData.command = $command
    $operationData.references = Get-ReferenceSnapshot
    $changes += [pscustomobject]@{ kind = 'reference'; target = $command; detail = '已执行书签、脚注、题注或交叉引用操作' }
  }
`;
}
