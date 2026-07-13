const WORD_REVISION_OPERATIONS = new Set([
  "inspectRevisions",
  "manageRevisions",
  "compareDocuments",
  "applyTrackedChanges",
]);

export function buildWordRevisionOperationScript(operation: string): string | undefined {
  if (!WORD_REVISION_OPERATIONS.has(operation)) return undefined;
  return String.raw`
  function Get-RevisionSnapshot() {
    $revisions = @()
    foreach ($revision in $doc.Revisions) {
      $revisions += [pscustomobject]@{
        index = [int]$revision.Index
        type = [int]$revision.Type
        author = [string]$revision.Author
        date = $(try { ([DateTime]$revision.Date).ToString('o') } catch { '' })
        text = $(try { [string]$revision.Range.Text } catch { '' })
        start = $(try { [int]$revision.Range.Start } catch { -1 })
        end = $(try { [int]$revision.Range.End } catch { -1 })
      }
    }
    $comments = @()
    foreach ($comment in $doc.Comments) {
      $comments += [pscustomobject]@{
        index = [int]$comment.Index
        author = [string]$comment.Author
        initials = [string]$comment.Initial
        date = $(try { ([DateTime]$comment.Date).ToString('o') } catch { '' })
        text = [string]$comment.Range.Text
        scope = $(try { [string]$comment.Scope.Text } catch { '' })
        start = $(try { [int]$comment.Scope.Start } catch { -1 })
        end = $(try { [int]$comment.Scope.End } catch { -1 })
      }
    }
    return [pscustomobject]@{
      trackRevisions = [bool]$doc.TrackRevisions
      revisionCount = [int]$doc.Revisions.Count
      commentCount = [int]$doc.Comments.Count
      revisions = $revisions
      comments = $comments
    }
  }
  function Test-RevisionRule($revision, $rule) {
    if ($null -eq $rule) { return $true }
    $authors = @($rule.authors | ForEach-Object { [string]$_ })
    if ($authors.Count -gt 0 -and [string]$revision.Author -notin $authors) { return $false }
    $types = @($rule.types | ForEach-Object { [int]$_ })
    if ($types.Count -gt 0 -and [int]$revision.Type -notin $types) { return $false }
    if ($rule.start -ne $null -and [int]$revision.Range.End -lt [int]$rule.start) { return $false }
    if ($rule.end -ne $null -and [int]$revision.Range.Start -gt [int]$rule.end) { return $false }
    if ($rule.textPattern -and -not [regex]::IsMatch([string]$revision.Range.Text, [string]$rule.textPattern)) { return $false }
    if ($rule.fromDate) { try { if ([DateTime]$revision.Date -lt [DateTime]::Parse([string]$rule.fromDate)) { return $false } } catch {} }
    if ($rule.toDate) { try { if ([DateTime]$revision.Date -gt [DateTime]::Parse([string]$rule.toDate)) { return $false } } catch {} }
    return $true
  }
  function Test-CommentRule($comment, $rule) {
    if ($null -eq $rule) { return $true }
    $authors = @($rule.authors | ForEach-Object { [string]$_ })
    if ($authors.Count -gt 0 -and [string]$comment.Author -notin $authors) { return $false }
    if ($rule.textPattern -and -not [regex]::IsMatch([string]$comment.Range.Text, [string]$rule.textPattern)) { return $false }
    return $true
  }
  function Apply-TrackedFindEdit($edit) {
    if (-not $edit.find) { throw 'replace/delete 修订需要 edit.find' }
    $count = 0
    $searchStart = 0
    $limit = [Math]::Max(1, $(if ($edit.maxReplacements) { [int]$edit.maxReplacements } else { 1000 }))
    $seenRanges = @{}
    while ($searchStart -lt $doc.Content.End -and $count -lt $limit) {
      $search = $doc.Range($searchStart, $doc.Content.End - 1)
      $find = $search.Find
      $find.ClearFormatting()
      $find.Text = [string]$edit.find
      $find.Forward = $true
      $find.Wrap = 0
      $find.MatchCase = [bool]$edit.matchCase
      $find.MatchWildcards = [bool]$edit.matchWildcards
      if (-not $find.Execute()) { break }
      $matched = $search.Duplicate
      $originalStart = [int]$matched.Start
      $originalEnd = [int]$matched.End
      $rangeKey = $originalStart.ToString() + ':' + $originalEnd.ToString()
      if ($seenRanges.ContainsKey($rangeKey)) { break }
      $seenRanges[$rangeKey] = $true
      $matched.Text = if ([string]$edit.command -eq 'delete') { '' } else { [string]$edit.replace }
      $count++
      $searchStart = [Math]::Max($originalEnd + 1, $originalStart + ([string]$edit.replace).Length + 1)
      if ($edit.all -eq $false) { break }
    }
    return $count
  }
  function Apply-TrackedEdit($edit) {
    $command = if ($edit.command) { [string]$edit.command } else { 'replace' }
    if ($command -in @('replace', 'delete')) { return Apply-TrackedFindEdit $edit }
    if ($command -eq 'insert') {
      $position = if ($edit.position -eq 'start') { 0 } elseif ($edit.position -eq 'end' -or $null -eq $edit.position) { [Math]::Max(0, $doc.Content.End - 1) } else { [Math]::Max(0, [Math]::Min([int]$edit.position, $doc.Content.End - 1)) }
      $doc.Range($position, $position).InsertAfter([string]$edit.text)
      return 1
    }
    if ($command -eq 'replaceBookmark') {
      if (-not $edit.name -or -not $doc.Bookmarks.Exists([string]$edit.name)) { throw 'replaceBookmark 找不到书签: ' + [string]$edit.name }
      $target = $doc.Bookmarks.Item([string]$edit.name).Range
      $start = $target.Start
      $target.Text = [string]$edit.text
      [void]$doc.Bookmarks.Add([string]$edit.name, $doc.Range($start, $start + ([string]$edit.text).Length))
      return 1
    }
    if ($command -eq 'replaceContentControl') {
      $count = 0
      foreach ($control in $doc.ContentControls) {
        if (($edit.tag -and [string]$control.Tag -eq [string]$edit.tag) -or ($edit.title -and [string]$control.Title -eq [string]$edit.title)) {
          try { $control.LockContents = $false } catch {}
          $control.Range.Text = [string]$edit.text
          $count++
        }
      }
      return $count
    }
    throw '不支持的修订编辑命令: ' + $command
  }
  function Get-DocumentParagraphTexts($targetDocument) {
    $items = @()
    foreach ($paragraph in $targetDocument.Paragraphs) {
      $text = ([string]$paragraph.Range.Text).Trim([char]13, [char]7).Trim()
      if ($text) { $items += $text }
    }
    return $items
  }
  function Compare-ParagraphTexts($original, $revised) {
    $originalCount = $original.Count
    $revisedCount = $revised.Count
    $maxLcsParagraphs = if ($actionParams.maxLcsParagraphs) { [Math]::Max(50, [int]$actionParams.maxLcsParagraphs) } else { 500 }
    $changes = @()
    if ($originalCount -le $maxLcsParagraphs -and $revisedCount -le $maxLcsParagraphs) {
      $matrix = New-Object 'int[,]' ($originalCount + 1), ($revisedCount + 1)
      for ($i = $originalCount - 1; $i -ge 0; $i--) {
        for ($j = $revisedCount - 1; $j -ge 0; $j--) {
          $nextI = $i + 1
          $nextJ = $j + 1
          if ([string]$original[$i] -ceq [string]$revised[$j]) {
            $diagonal = $matrix[$nextI,$nextJ]
            $matrix[$i,$j] = $diagonal + 1
          } else {
            $down = $matrix[$nextI,$j]
            $right = $matrix[$i,$nextJ]
            $matrix[$i,$j] = [Math]::Max($down, $right)
          }
        }
      }
      $i = 0; $j = 0
      while ($i -lt $originalCount -or $j -lt $revisedCount) {
        if ($i -lt $originalCount -and $j -lt $revisedCount -and [string]$original[$i] -ceq [string]$revised[$j]) { $i++; $j++; continue }
        $nextI = $i + 1
        $nextJ = $j + 1
        $right = if ($i -lt $originalCount -and $j -lt $revisedCount) { $matrix[$i,$nextJ] } else { 0 }
        $down = if ($i -lt $originalCount -and $j -lt $revisedCount) { $matrix[$nextI,$j] } else { 0 }
        if ($j -lt $revisedCount -and ($i -ge $originalCount -or $right -ge $down)) {
          $changes += [pscustomobject]@{ type = 'added'; originalIndex = $null; revisedIndex = $j + 1; text = [string]$revised[$j] }
          $j++
        } else {
          $changes += [pscustomobject]@{ type = 'deleted'; originalIndex = $i + 1; revisedIndex = $null; text = [string]$original[$i] }
          $i++
        }
      }
    } else {
      $max = [Math]::Max($originalCount, $revisedCount)
      for ($index = 0; $index -lt $max; $index++) {
        $before = if ($index -lt $originalCount) { [string]$original[$index] } else { '' }
        $after = if ($index -lt $revisedCount) { [string]$revised[$index] } else { '' }
        if ($before -ceq $after) { continue }
        if ($before) { $changes += [pscustomobject]@{ type = 'deleted'; originalIndex = $index + 1; revisedIndex = $null; text = $before } }
        if ($after) { $changes += [pscustomobject]@{ type = 'added'; originalIndex = $null; revisedIndex = $index + 1; text = $after } }
      }
    }
    return $changes
  }
  function Write-ComparisonReport($comparisonDocument, $changes, [string]$originalName, [string]$revisedName) {
    $comparisonDocument.Content.Text = '文档对比摘要' + [Environment]::NewLine +
      '原文：' + $originalName + [Environment]::NewLine +
      '修订稿：' + $revisedName + [Environment]::NewLine +
      '新增段落：' + @($changes | Where-Object { $_.type -eq 'added' }).Count + '；删除段落：' + @($changes | Where-Object { $_.type -eq 'deleted' }).Count + [Environment]::NewLine + [Environment]::NewLine
    try { $comparisonDocument.Paragraphs.Item(1).Range.Style = $comparisonDocument.Styles.Item(-2) } catch {}
    $limit = if ($actionParams.maxReportChanges) { [Math]::Max(1, [int]$actionParams.maxReportChanges) } else { 500 }
    $written = 0
    foreach ($change in $changes) {
      if ($written -ge $limit) { break }
      $prefix = if ($change.type -eq 'added') { '[新增] ' } else { '[删除] ' }
      $start = [Math]::Max(0, $comparisonDocument.Content.End - 1)
      $comparisonDocument.Range($start, $start).InsertAfter($prefix + [string]$change.text + [Environment]::NewLine)
      $line = $comparisonDocument.Range($start, [Math]::Max($start, $comparisonDocument.Content.End - 1))
      if ($change.type -eq 'added') { $line.Font.Color = 5287936; $line.Font.Underline = 1 }
      else { $line.Font.Color = 255; $line.Font.StrikeThrough = $true }
      $written++
    }
    return $written
  }

  if ($_operation -eq 'inspectRevisions') {
    $operationData.review = Get-RevisionSnapshot
    $changes = @()
  } elseif ($_operation -eq 'compareDocuments') {
    if (-not $actionParams.revisedFilePath) { throw 'compareDocuments 需要 params.revisedFilePath' }
    $revised = $null; $compared = $null
    try {
      $revised = $app.Documents.Open([string]$actionParams.revisedFilePath, $false, $true)
      $originalParagraphs = @(Get-DocumentParagraphTexts $doc)
      $revisedParagraphs = @(Get-DocumentParagraphTexts $revised)
      $diff = @(Compare-ParagraphTexts $originalParagraphs $revisedParagraphs)
      $compared = $app.Documents.Add()
      $written = Write-ComparisonReport $compared $diff ([string]$doc.Name) ([string]$revised.Name)
      $compared.SaveAs2($_outputPath, 16)
      $operationData.summary = [pscustomobject]@{
        originalParagraphs = $originalParagraphs.Count
        revisedParagraphs = $revisedParagraphs.Count
        changeCount = $diff.Count
        addedCount = @($diff | Where-Object { $_.type -eq 'added' }).Count
        deletedCount = @($diff | Where-Object { $_.type -eq 'deleted' }).Count
        reportChanges = $written
        changes = @($diff | Select-Object -First 100)
      }
      $changes += [pscustomobject]@{ kind = 'document-compare'; target = $_outputPath; detail = '已对比两个文档并生成段落级差异报告' }
    } finally {
      if ($null -ne $compared) { try { $compared.Close($false) } catch {} }
      if ($null -ne $revised) { try { $revised.Close($false) } catch {} }
      if ($null -ne $compared) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($compared) } catch {}; $compared = $null }
      if ($null -ne $revised) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($revised) } catch {}; $revised = $null }
    }
  } elseif ($_operation -eq 'applyTrackedChanges') {
    if ($null -eq $actionParams.edits) { throw 'applyTrackedChanges 需要 params.edits 数组' }
    $wasTracking = [bool]$doc.TrackRevisions
    $doc.TrackRevisions = $true
    $applied = @()
    foreach ($edit in @($actionParams.edits)) {
      $count = Apply-TrackedEdit $edit
      $applied += [pscustomobject]@{ command = $(if ($edit.command) { [string]$edit.command } else { 'replace' }); count = $count }
    }
    if ($actionParams.keepTracking -eq $false) { $doc.TrackRevisions = $wasTracking }
    $operationData.applied = $applied
    $operationData.review = Get-RevisionSnapshot
    $changes += [pscustomobject]@{ kind = 'revision'; target = 'document'; detail = '已在修订模式下应用 AI 修改并保留原文轨迹' }
  } else {
    $command = if ($actionParams.command) { [string]$actionParams.command } else { 'track' }
    $processed = 0
    switch ($command) {
      'track' { $doc.TrackRevisions = $actionParams.enabled -ne $false }
      'acceptAll' { $processed = [int]$doc.Revisions.Count; if ($processed -gt 0) { $doc.AcceptAllRevisions() } }
      'rejectAll' { $processed = [int]$doc.Revisions.Count; if ($processed -gt 0) { $doc.RejectAllRevisions() } }
      'acceptMatching' {
        for ($index = $doc.Revisions.Count; $index -ge 1; $index--) { $revision = $doc.Revisions.Item($index); if (Test-RevisionRule $revision $actionParams.rule) { $revision.Accept(); $processed++ } }
      }
      'rejectMatching' {
        for ($index = $doc.Revisions.Count; $index -ge 1; $index--) { $revision = $doc.Revisions.Item($index); if (Test-RevisionRule $revision $actionParams.rule) { $revision.Reject(); $processed++ } }
      }
      'deleteComments' { $processed = [int]$doc.Comments.Count; while ($doc.Comments.Count -gt 0) { $doc.Comments.Item(1).Delete() } }
      'deleteCommentsMatching' {
        for ($index = $doc.Comments.Count; $index -ge 1; $index--) { $comment = $doc.Comments.Item($index); if (Test-CommentRule $comment $actionParams.rule) { $comment.Delete(); $processed++ } }
      }
      default { throw '不支持的审阅命令: ' + $command }
    }
    $operationData.command = $command
    $operationData.processed = $processed
    $operationData.review = Get-RevisionSnapshot
    $changes += [pscustomobject]@{ kind = 'revision'; target = $command; detail = '已按规则处理修订或批注' }
  }
`;
}
