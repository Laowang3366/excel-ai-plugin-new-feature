const WORD_MAIL_MERGE_OPERATIONS = new Set([
  "prepareMailMergeTemplate",
  "mailMerge",
  "batchMailMerge",
]);

export function buildWordMailMergeOperationScript(operation: string): string | undefined {
  if (!WORD_MAIL_MERGE_OPERATIONS.has(operation)) return undefined;
  return String.raw`
  function Replace-WordText($targetDocument, [string]$findText, [string]$replacement) {
    if (-not $findText) { return 0 }
    $count = 0; $start = 0
    while ($start -lt $targetDocument.Content.End) {
      $search = $targetDocument.Range($start, $targetDocument.Content.End - 1)
      $finder = $search.Find
      $finder.ClearFormatting()
      $finder.Text = $findText
      $finder.Forward = $true
      $finder.Wrap = 0
      if (-not $finder.Execute()) { break }
      $matched = $search.Duplicate
      $matched.Text = $replacement
      $count++
      $start = [Math]::Max($matched.End, $matched.Start + 1)
    }
    return $count
  }
  function Get-MergeRecordValues($dataSource) {
    $values = @{}
    foreach ($field in $dataSource.DataFields) {
      try { $values[[string]$field.Name] = [string]$field.Value } catch {}
    }
    return $values
  }
  function Test-MailMergeCondition([string]$actual, $rule) {
    $expected = [string]$rule.value
    switch ([string]$rule.operator) {
      'ne' { return $actual -ne $expected }
      'contains' { return $actual.Contains($expected) }
      'notContains' { return -not $actual.Contains($expected) }
      'empty' { return [string]::IsNullOrWhiteSpace($actual) }
      'notEmpty' { return -not [string]::IsNullOrWhiteSpace($actual) }
      'gt' { return [double]$actual -gt [double]$expected }
      'gte' { return [double]$actual -ge [double]$expected }
      'lt' { return [double]$actual -lt [double]$expected }
      'lte' { return [double]$actual -le [double]$expected }
      default { return $actual -eq $expected }
    }
  }
  function Apply-MailMergeConditions($targetDocument, $recordValues) {
    foreach ($rule in @($actionParams.conditions)) {
      if (-not $rule.placeholder -or -not $rule.field) { continue }
      $actual = if ($recordValues.ContainsKey([string]$rule.field)) { [string]$recordValues[[string]$rule.field] } else { '' }
      $replacement = if (Test-MailMergeCondition $actual $rule) { [string]$rule.trueText } else { [string]$rule.falseText }
      [void](Replace-WordText $targetDocument ([string]$rule.placeholder) $replacement)
    }
  }
  function Apply-MailMergeImages($targetDocument, $recordValues) {
    foreach ($imageRule in @($actionParams.imageFields)) {
      if (-not $imageRule.placeholder -or -not $imageRule.field) { continue }
      $imagePath = if ($recordValues.ContainsKey([string]$imageRule.field)) { [string]$recordValues[[string]$imageRule.field] } else { '' }
      $search = $targetDocument.Content.Duplicate
      $finder = $search.Find
      $finder.ClearFormatting(); $finder.Text = [string]$imageRule.placeholder; $finder.Forward = $true; $finder.Wrap = 0
      if (-not $finder.Execute()) { continue }
      if (-not $imagePath -or -not [IO.File]::Exists($imagePath)) {
        if ($actionParams.strictImages -eq $true) { throw '图片字段文件不存在: ' + $imagePath }
        $search.Text = ''
        continue
      }
      $search.Text = ''
      $picture = $targetDocument.InlineShapes.AddPicture($imagePath, $false, $true, $search)
      if ($imageRule.width) { $picture.LockAspectRatio = -1; $picture.Width = [double]$imageRule.width }
      if ($imageRule.height) { $picture.LockAspectRatio = -1; $picture.Height = [double]$imageRule.height }
    }
  }
  function Get-SafeOutputName($recordValues, [int]$recordIndex) {
    $name = ''
    if ($actionParams.fileNamePattern) {
      $name = [string]$actionParams.fileNamePattern
      foreach ($key in $recordValues.Keys) { $name = $name.Replace('{' + [string]$key + '}', [string]$recordValues[$key]) }
    } elseif ($actionParams.fileNameField -and $recordValues.ContainsKey([string]$actionParams.fileNameField)) {
      $name = [string]$recordValues[[string]$actionParams.fileNameField]
    }
    if ([string]::IsNullOrWhiteSpace($name)) { $name = 'record-' + $recordIndex.ToString('0000') }
    $name = $name -replace '[\/:*?"<>|]', '_'
    $name = $name.Trim().TrimEnd('.')
    if (-not $name) { return 'record-' + $recordIndex.ToString('0000') }
    return $name
  }
  function Open-MailMergeDataSource($merge) {
    $sheetName = if ($actionParams.dataSheetName) { [string]$actionParams.dataSheetName } else { 'Sheet1' }
    $safeSheetName = $sheetName.Replace(']', ']]')
    $sql = 'SELECT * FROM [' + $safeSheetName + '$]'
    $merge.OpenDataSource(
      [string]$actionParams.dataSourcePath,
      0, $false, $true, $true, $false,
      '', '', $false, '', '', '',
      $sql, '', $false, 0
    )
  }
  function Save-MergedRecord($mergedDocument, [string]$outputDirectory, [string]$baseName) {
    $format = if ($actionParams.outputFormat) { [string]$actionParams.outputFormat } else { 'docx' }
    $outputs = @()
    if ($format -in @('docx', 'both')) {
      $docxPath = [IO.Path]::Combine($outputDirectory, $baseName + '.docx')
      if ([IO.File]::Exists($docxPath) -and $actionParams.overwrite -ne $true) { throw '输出文件已存在，请设置 overwrite=true: ' + $docxPath }
      $mergedDocument.SaveAs2($docxPath, 16)
      $outputs += $docxPath
    }
    if ($format -in @('pdf', 'both')) {
      $pdfPath = [IO.Path]::Combine($outputDirectory, $baseName + '.pdf')
      if ([IO.File]::Exists($pdfPath) -and $actionParams.overwrite -ne $true) { throw '输出文件已存在，请设置 overwrite=true: ' + $pdfPath }
      $mergedDocument.ExportAsFixedFormat($pdfPath, 17)
      $outputs += $pdfPath
    }
    return $outputs
  }
  function Prepare-MailMergeFields() {
    $created = @()
    foreach ($fieldRule in @($actionParams.fields)) {
      if (-not $fieldRule.placeholder -or -not $fieldRule.field) { continue }
      $start = 0
      while ($start -lt $doc.Content.End) {
        $search = $doc.Range($start, $doc.Content.End - 1)
        $finder = $search.Find
        $finder.ClearFormatting(); $finder.Text = [string]$fieldRule.placeholder; $finder.Forward = $true; $finder.Wrap = 0
        if (-not $finder.Execute()) { break }
        $matched = $search.Duplicate
        $position = $matched.Start
        $matched.Text = ''
        $fieldRange = $doc.Range($position, $position)
        [void]$doc.Fields.Add($fieldRange, 59, [string]$fieldRule.field, $true)
        $created += [pscustomobject]@{ placeholder = [string]$fieldRule.placeholder; field = [string]$fieldRule.field; position = $position }
        $start = $fieldRange.End + 1
        if ($fieldRule.all -eq $false) { break }
      }
    }
    return $created
  }

  if ($_operation -eq 'prepareMailMergeTemplate') {
    $created = @(Prepare-MailMergeFields)
    $operationData.createdFields = $created
    $operationData.createdCount = $created.Count
    $changes += [pscustomobject]@{ kind = 'mail-merge-template'; target = 'document'; detail = '已将模板占位符转换为邮件合并域' }
  } elseif ($_operation -eq 'mailMerge' -and $actionParams.mode -ne 'separate') {
    if (-not $actionParams.dataSourcePath) { throw 'mailMerge 需要 params.dataSourcePath' }
    $merge = $doc.MailMerge
    Open-MailMergeDataSource $merge
    $dataSource = $merge.DataSource
    $recordCount = [int]$dataSource.RecordCount
    $first = if ($actionParams.firstRecord) { [Math]::Max(1, [int]$actionParams.firstRecord) } else { 1 }
    $last = if ($actionParams.lastRecord) { [Math]::Min($recordCount, [int]$actionParams.lastRecord) } else { $recordCount }
    $stagingDirectory = [IO.Path]::Combine([IO.Path]::GetTempPath(), 'wengge-mail-merge-' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($stagingDirectory)
    $stagedDocuments = @()
    $mergedRecords = 0
    try {
      for ($recordIndex = $first; $recordIndex -le $last; $recordIndex++) {
        $doc.Activate()
        $dataSource.ActiveRecord = $recordIndex
        $recordValues = Get-MergeRecordValues $dataSource
        $dataSource.FirstRecord = $recordIndex
        $dataSource.LastRecord = $recordIndex
        $merge.Destination = 0
        $merge.Execute($false)
        $mergedDocument = $app.ActiveDocument
        try {
          Apply-MailMergeConditions $mergedDocument $recordValues
          Apply-MailMergeImages $mergedDocument $recordValues
          foreach ($field in $mergedDocument.Fields) { try { $field.Update() } catch {} }
          $stagedPath = [IO.Path]::Combine($stagingDirectory, ('record-{0:D6}.docx' -f $recordIndex))
          $mergedDocument.SaveAs2($stagedPath, 16)
          $stagedDocuments += $stagedPath
          $mergedRecords++
        } finally {
          if ($null -ne $mergedDocument) { try { $mergedDocument.Close($false) } catch {}; try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($mergedDocument) } catch {}; $mergedDocument = $null }
        }
      }
      $combinedDocument = $app.Documents.Add()
      try {
        for ($stageIndex = 0; $stageIndex -lt $stagedDocuments.Count; $stageIndex++) {
          $insertAt = [Math]::Max(0, $combinedDocument.Content.End - 1)
          $insertRange = $combinedDocument.Range($insertAt, $insertAt)
          if ($stageIndex -gt 0) {
            $insertRange.InsertBreak(7)
            $insertAt = [Math]::Max(0, $combinedDocument.Content.End - 1)
            $insertRange = $combinedDocument.Range($insertAt, $insertAt)
          }
          $insertRange.InsertFile([string]$stagedDocuments[$stageIndex])
        }
        $format = if ($actionParams.outputFormat) { [string]$actionParams.outputFormat } elseif ([IO.Path]::GetExtension($_outputPath) -eq '.pdf') { 'pdf' } else { 'docx' }
        if ($format -eq 'pdf') { $combinedDocument.ExportAsFixedFormat($_outputPath, 17) }
        else { $combinedDocument.SaveAs2($_outputPath, 16) }
      } finally {
        if ($null -ne $combinedDocument) { try { $combinedDocument.Close($false) } catch {}; try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($combinedDocument) } catch {}; $combinedDocument = $null }
      }
      $operationData.recordCount = $mergedRecords
      $operationData.outputPaths = @($_outputPath)
    } finally {
      if ([IO.Directory]::Exists($stagingDirectory)) { try { [IO.Directory]::Delete($stagingDirectory, $true) } catch {} }
    }
    $changes += [pscustomobject]@{ kind = 'mail-merge'; target = $_outputPath; detail = '已生成合并文档' }
  } else {
    if (-not $actionParams.dataSourcePath) { throw 'batchMailMerge 需要 params.dataSourcePath' }
    $outputDirectory = if ($actionParams.outputDirectory) { [string]$actionParams.outputDirectory } else { [IO.Path]::ChangeExtension($_outputPath, $null) }
    if (-not [IO.Directory]::Exists($outputDirectory)) { [void][IO.Directory]::CreateDirectory($outputDirectory) }
    $merge = $doc.MailMerge
    Open-MailMergeDataSource $merge
    $dataSource = $merge.DataSource
    $recordCount = [int]$dataSource.RecordCount
    $first = if ($actionParams.firstRecord) { [Math]::Max(1, [int]$actionParams.firstRecord) } else { 1 }
    $last = if ($actionParams.lastRecord) { [Math]::Min($recordCount, [int]$actionParams.lastRecord) } else { $recordCount }
    $records = @(); $allOutputs = @()
    for ($recordIndex = $first; $recordIndex -le $last; $recordIndex++) {
      $doc.Activate()
      $dataSource.ActiveRecord = $recordIndex
      $recordValues = Get-MergeRecordValues $dataSource
      $baseName = Get-SafeOutputName $recordValues $recordIndex
      $dataSource.FirstRecord = $recordIndex
      $dataSource.LastRecord = $recordIndex
      $merge.Destination = 0
      $merge.Execute($false)
      $mergedDocument = $app.ActiveDocument
      try {
        Apply-MailMergeConditions $mergedDocument $recordValues
        Apply-MailMergeImages $mergedDocument $recordValues
        foreach ($field in $mergedDocument.Fields) { try { $field.Update() } catch {} }
        $outputs = @(Save-MergedRecord $mergedDocument $outputDirectory $baseName)
        $allOutputs += $outputs
        $records += [pscustomobject]@{ record = $recordIndex; name = $baseName; outputPaths = $outputs }
      } finally { $mergedDocument.Close($false) }
    }
    $_outputPath = $outputDirectory
    $operationData.recordCount = $records.Count
    $operationData.records = $records
    $operationData.outputPaths = $allOutputs
    $changes += [pscustomobject]@{ kind = 'mail-merge'; target = $outputDirectory; detail = '已按记录批量生成并命名文档' }
  }
`;
}
