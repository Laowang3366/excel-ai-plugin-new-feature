const PRESENTATION_PLAYBACK_OPERATIONS = new Set([
  "inspectAnimations",
  "configureAnimations",
  "configureSlideShow",
  "setSpeakerNotes",
  "inspectSpeakerNotes",
  "exportHandouts",
]);

export function buildPresentationPlaybackOperationScript(operation: string): string | undefined {
  if (!PRESENTATION_PLAYBACK_OPERATIONS.has(operation)) return undefined;
  return String.raw`
  function Resolve-AnimationEffect([string]$name, [string]$category) {
    if ($category -eq 'emphasis') {
      switch ($name) { 'growShrink' { 59 } 'spin' { 61 } 'transparency' { 62 } default { 59 } }
      return
    }
    switch ($name) {
      'appear' { 1 }
      'fly' { 2 }
      'dissolve' { 9 }
      'fade' { 10 }
      'wipe' { 22 }
      'zoom' { 23 }
      default { 10 }
    }
  }
  function Resolve-AnimationTrigger([string]$trigger) {
    switch ($trigger) { 'withPrevious' { 2 } 'afterPrevious' { 3 } default { 1 } }
  }
  function Get-AnimationSnapshot($targetSlide) {
    $effects = @()
    $sequence = $targetSlide.TimeLine.MainSequence
    for ($effectIndex = 1; $effectIndex -le $sequence.Count; $effectIndex++) {
      $effect = $sequence.Item($effectIndex)
      $behaviors = @()
      try {
        for ($behaviorIndex = 1; $behaviorIndex -le $effect.Behaviors.Count; $behaviorIndex++) {
          $behavior = $effect.Behaviors.Item($behaviorIndex)
          $behaviors += [pscustomobject]@{
            type = [int]$behavior.Type
            byX = $(try { [double]$behavior.MotionEffect.ByX } catch { 0 })
            byY = $(try { [double]$behavior.MotionEffect.ByY } catch { 0 })
          }
        }
      } catch {}
      $effects += [pscustomobject]@{
        index = $effectIndex
        shapeName = $(try { [string]$effect.Shape.Name } catch { '' })
        effectType = $(try { [int]$effect.EffectType } catch { 0 })
        exit = $(try { [bool]$effect.Exit } catch { $false })
        trigger = $(try { [int]$effect.Timing.TriggerType } catch { 0 })
        duration = $(try { [double]$effect.Timing.Duration } catch { 0 })
        delay = $(try { [double]$effect.Timing.TriggerDelayTime } catch { 0 })
        repeatCount = $(try { [double]$effect.Timing.RepeatCount } catch { 0 })
        behaviors = $behaviors
      }
    }
    $transition = $targetSlide.SlideShowTransition
    return [pscustomobject]@{
      slideIndex = [int]$targetSlide.SlideIndex
      effects = $effects
      transition = [pscustomobject]@{
        entryEffect = $(try { [int]$transition.EntryEffect } catch { 0 })
        advanceOnClick = $(try { [bool]$transition.AdvanceOnClick } catch { $false })
        advanceOnTime = $(try { [bool]$transition.AdvanceOnTime } catch { $false })
        advanceTime = $(try { [double]$transition.AdvanceTime } catch { 0 })
        duration = $(try { [double]$transition.Duration } catch { 0 })
      }
    }
  }
  function Get-SlideNotesText($targetSlide) {
    foreach ($shape in $targetSlide.NotesPage.Shapes) {
      try { if ($shape.PlaceholderFormat.Type -eq 2) { return [string]$shape.TextFrame.TextRange.Text } } catch {}
    }
    return ''
  }
  function Set-SlideNotesText($targetSlide, [string]$text, [bool]$append) {
    foreach ($shape in $targetSlide.NotesPage.Shapes) {
      try {
        if ($shape.PlaceholderFormat.Type -eq 2) {
          if ($append -and $shape.TextFrame.TextRange.Text) { $shape.TextFrame.TextRange.Text = ([string]$shape.TextFrame.TextRange.Text).Trim() + [Environment]::NewLine + $text }
          else { $shape.TextFrame.TextRange.Text = $text }
          return $true
        }
      } catch {}
    }
    return $false
  }
  function Get-SlideVisibleText($targetSlide) {
    $parts = @()
    foreach ($shape in $targetSlide.Shapes) { try { if ($shape.HasTextFrame -and $shape.TextFrame.HasText) { $parts += [string]$shape.TextFrame.TextRange.Text } } catch {} }
    return ($parts -join [Environment]::NewLine)
  }
  function Get-NotesCorrespondenceScore([string]$slideText, [string]$notesText) {
    if ([string]::IsNullOrWhiteSpace($slideText) -or [string]::IsNullOrWhiteSpace($notesText)) { return 0 }
    $slideTerms = @([regex]::Matches($slideText.ToLowerInvariant(), '[\p{L}\p{Nd}]') | ForEach-Object { $_.Value } | Sort-Object -Unique)
    $noteTerms = @([regex]::Matches($notesText.ToLowerInvariant(), '[\p{L}\p{Nd}]') | ForEach-Object { $_.Value } | Sort-Object -Unique)
    if ($slideTerms.Count -eq 0 -or $noteTerms.Count -eq 0) { return 0 }
    $common = @($noteTerms | Where-Object { $slideTerms -contains $_ }).Count
    return [Math]::Round($common / [Math]::Min($slideTerms.Count, $noteTerms.Count), 3)
  }
  function Read-PresentationZipText($archive, [string]$entryName) {
    $entry = $archive.GetEntry($entryName)
    if ($null -eq $entry) { return $null }
    $reader = [IO.StreamReader]::new($entry.Open(), [Text.Encoding]::UTF8, $true)
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
  }
  function Write-PresentationZipXml($archive, [string]$entryName, [Xml.XmlDocument]$xml) {
    $existing = $archive.GetEntry($entryName); if ($null -ne $existing) { $existing.Delete() }
    $entry = $archive.CreateEntry($entryName, [IO.Compression.CompressionLevel]::Optimal)
    $settings = [Xml.XmlWriterSettings]::new(); $settings.Encoding = [Text.UTF8Encoding]::new($false); $settings.Indent = $false
    $stream = $entry.Open(); $writer = [Xml.XmlWriter]::Create($stream, $settings)
    try { $xml.Save($writer) } finally { $writer.Dispose(); $stream.Dispose() }
  }
  function ConvertTo-PresentationXml([string]$text) {
    $xml = [Xml.XmlDocument]::new(); $xml.PreserveWhitespace = $true; $xml.LoadXml($text); return $xml
  }
  function Set-PresentationNotesXmlText([Xml.XmlDocument]$xml, [string]$text, [bool]$append) {
    $ns = [Xml.XmlNamespaceManager]::new($xml.NameTable)
    $ns.AddNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main')
    $ns.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
    $bodyPlaceholder = $xml.SelectSingleNode('//p:sp[p:nvSpPr/p:nvPr/p:ph[@type="body"]]', $ns)
    if ($null -eq $bodyPlaceholder) { throw '备注页模板缺少正文占位符' }
    $textBody = $bodyPlaceholder.SelectSingleNode('p:txBody', $ns)
    if ($null -eq $textBody) { throw '备注页模板缺少文本容器' }
    $existingText = @($textBody.SelectNodes('.//a:t', $ns) | ForEach-Object { $_.InnerText }) -join ''
    $finalText = if ($append -and -not [string]::IsNullOrWhiteSpace($existingText)) { $existingText.Trim() + [Environment]::NewLine + $text } else { $text }
    foreach ($paragraph in @($textBody.SelectNodes('a:p', $ns))) { [void]$textBody.RemoveChild($paragraph) }
    $paragraph = $xml.CreateElement('a', 'p', 'http://schemas.openxmlformats.org/drawingml/2006/main')
    $run = $xml.CreateElement('a', 'r', 'http://schemas.openxmlformats.org/drawingml/2006/main')
    $runProperties = $xml.CreateElement('a', 'rPr', 'http://schemas.openxmlformats.org/drawingml/2006/main'); $runProperties.SetAttribute('lang', 'zh-CN')
    $textNode = $xml.CreateElement('a', 't', 'http://schemas.openxmlformats.org/drawingml/2006/main'); $textNode.InnerText = $finalText
    [void]$run.AppendChild($runProperties); [void]$run.AppendChild($textNode); [void]$paragraph.AppendChild($run); [void]$textBody.AppendChild($paragraph)
  }
  function Update-PresentationNotesPackage([string]$path, $updates) {
    if (-not [IO.File]::Exists($path)) { throw '备注回退更新找不到演示文件: ' + $path }
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = $null; $lastError = $null
    foreach ($attempt in 1..4) {
      try { $archive = [IO.Compression.ZipFile]::Open($path, [IO.Compression.ZipArchiveMode]::Update); break }
      catch { $lastError = $_; Start-Sleep -Milliseconds (250 * $attempt) }
    }
    if ($null -eq $archive) { throw $lastError }
    try {
      $notesEntries = @($archive.Entries | Where-Object { $_.FullName -match '^ppt/notesSlides/notesSlide[0-9]+[.]xml$' } | Sort-Object FullName)
      if ($notesEntries.Count -eq 0) { throw 'WPS 未生成可复用的备注页模板' }
      $templateXmlText = Read-PresentationZipText $archive $notesEntries[0].FullName
      $templateRelsName = 'ppt/notesSlides/_rels/' + [IO.Path]::GetFileName($notesEntries[0].FullName) + '.rels'
      $templateRelsText = Read-PresentationZipText $archive $templateRelsName
      if (-not $templateRelsText) { throw 'WPS 备注页模板缺少关系文件' }
      $maxNotesIndex = @($notesEntries | ForEach-Object { if ($_.FullName -match 'notesSlide([0-9]+)[.]xml$') { [int]$Matches[1] } } | Measure-Object -Maximum).Maximum
      $contentTypes = ConvertTo-PresentationXml (Read-PresentationZipText $archive '[Content_Types].xml')
      $contentNs = [Xml.XmlNamespaceManager]::new($contentTypes.NameTable); $contentNs.AddNamespace('ct', 'http://schemas.openxmlformats.org/package/2006/content-types')
      foreach ($update in @($updates)) {
        $slideIndex = [int]$update.slideIndex
        $slideRelsName = 'ppt/slides/_rels/slide' + $slideIndex + '.xml.rels'
        $slideRelsText = Read-PresentationZipText $archive $slideRelsName
        if (-not $slideRelsText) { throw '找不到幻灯片关系文件: ' + $slideRelsName }
        $slideRels = ConvertTo-PresentationXml $slideRelsText
        $relsNs = [Xml.XmlNamespaceManager]::new($slideRels.NameTable); $relsNs.AddNamespace('r', 'http://schemas.openxmlformats.org/package/2006/relationships')
        $notesRelation = $slideRels.SelectSingleNode('/r:Relationships/r:Relationship[contains(@Type, "/notesSlide")]', $relsNs)
        if ($null -ne $notesRelation) {
          $notesFileName = [IO.Path]::GetFileName([string]$notesRelation.Target)
          $notesEntryName = 'ppt/notesSlides/' + $notesFileName
          $notesXml = ConvertTo-PresentationXml (Read-PresentationZipText $archive $notesEntryName)
        } else {
          $maxNotesIndex = [int]$maxNotesIndex + 1
          $notesFileName = 'notesSlide' + $maxNotesIndex + '.xml'
          $notesEntryName = 'ppt/notesSlides/' + $notesFileName
          $notesXml = ConvertTo-PresentationXml $templateXmlText
          $notesRels = ConvertTo-PresentationXml $templateRelsText
          $notesRelsNs = [Xml.XmlNamespaceManager]::new($notesRels.NameTable); $notesRelsNs.AddNamespace('r', 'http://schemas.openxmlformats.org/package/2006/relationships')
          $slideRelation = $notesRels.SelectSingleNode('/r:Relationships/r:Relationship[contains(@Type, "/slide") and not(contains(@Type, "/notesSlide"))]', $notesRelsNs)
          if ($null -eq $slideRelation) { throw '备注页模板缺少幻灯片关系' }
          $slideRelation.SetAttribute('Target', '../slides/slide' + $slideIndex + '.xml')
          Write-PresentationZipXml $archive ('ppt/notesSlides/_rels/' + $notesFileName + '.rels') $notesRels
          $usedIds = @($slideRels.SelectNodes('/r:Relationships/r:Relationship', $relsNs) | ForEach-Object { if ($_.Id -match '^rId([0-9]+)$') { [int]$Matches[1] } })
          $nextId = if ($usedIds.Count -gt 0) { [int](($usedIds | Measure-Object -Maximum).Maximum) + 1 } else { 1 }
          $notesRelation = $slideRels.CreateElement('Relationship', 'http://schemas.openxmlformats.org/package/2006/relationships')
          $notesRelation.SetAttribute('Id', 'rId' + $nextId); $notesRelation.SetAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide'); $notesRelation.SetAttribute('Target', '../notesSlides/' + $notesFileName)
          [void]$slideRels.DocumentElement.AppendChild($notesRelation)
          Write-PresentationZipXml $archive $slideRelsName $slideRels
          $partName = '/ppt/notesSlides/' + $notesFileName
          if ($null -eq $contentTypes.SelectSingleNode('/ct:Types/ct:Override[@PartName="' + $partName + '"]', $contentNs)) {
            $override = $contentTypes.CreateElement('Override', 'http://schemas.openxmlformats.org/package/2006/content-types')
            $override.SetAttribute('PartName', $partName); $override.SetAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml')
            [void]$contentTypes.DocumentElement.AppendChild($override)
          }
        }
        Set-PresentationNotesXmlText $notesXml ([string]$update.text) ([bool]$update.append)
        Write-PresentationZipXml $archive $notesEntryName $notesXml
      }
      Write-PresentationZipXml $archive '[Content_Types].xml' $contentTypes
    } finally { $archive.Dispose() }
  }

  if ($_operation -eq 'inspectAnimations') {
    $targetSlides = if ($actionParams.allSlides -eq $true) { @($pres.Slides) } else { @($slide) }
    $animationSlides = @(); foreach ($targetSlide in $targetSlides) { $animationSlides += Get-AnimationSnapshot $targetSlide }
    $operationData.animations = $animationSlides
    $operationData.slideShow = [pscustomobject]@{
      showType = $(try { [int]$pres.SlideShowSettings.ShowType } catch { 0 })
      advanceMode = $(try { [int]$pres.SlideShowSettings.AdvanceMode } catch { 0 })
      loopUntilStopped = $(try { [bool]$pres.SlideShowSettings.LoopUntilStopped } catch { $false })
      showWithAnimation = $(try { [bool]$pres.SlideShowSettings.ShowWithAnimation } catch { $false })
    }
    $changes = @()
  } elseif ($_operation -eq 'configureAnimations') {
    $sequence = $slide.TimeLine.MainSequence
    if ($actionParams.clearExisting -eq $true) { while ($sequence.Count -gt 0) { $sequence.Item(1).Delete() } }
    $rules = if ($actionParams.effects) { @($actionParams.effects) } else { @($actionParams) }
    $animated = @()
    foreach ($rule in $rules) {
      $category = if ($rule.category) { [string]$rule.category } else { 'entrance' }
      $shapeNames = @($rule.shapeNames | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
      if ($rule.shapeName) { $shapeNames += [string]$rule.shapeName }
      foreach ($shape in $slide.Shapes) {
        if ($shapeNames.Count -gt 0 -and $shapeNames -notcontains [string]$shape.Name) { continue }
        $effectId = Resolve-AnimationEffect ([string]$rule.effect) $category
        $trigger = Resolve-AnimationTrigger ([string]$rule.trigger)
        if ($rule.order) { $effect = $sequence.AddEffect($shape, $effectId, 0, $trigger, [int]$rule.order) }
        else { $effect = $sequence.AddEffect($shape, $effectId, 0, $trigger) }
        if ($category -eq 'exit') { $effect.Exit = -1 }
        if ($category -eq 'path') {
          $motionBehavior = $effect.Behaviors.Add(1)
          $motion = $motionBehavior.MotionEffect
          $motion.ByX = if ($null -ne $rule.pathX) { [single]$rule.pathX } else { 0.2 }
          $motion.ByY = if ($null -ne $rule.pathY) { [single]$rule.pathY } else { 0 }
        }
        if ($rule.duration) { $effect.Timing.Duration = [double]$rule.duration }
        if ($null -ne $rule.delay) { $effect.Timing.TriggerDelayTime = [double]$rule.delay }
        if ($rule.repeatCount) { try { $effect.Timing.RepeatCount = [double]$rule.repeatCount } catch {} }
        $animated += [pscustomobject]@{ shapeName = [string]$shape.Name; category = $category; effect = [string]$rule.effect }
      }
    }
    $operationData.animated = $animated
    $operationData.snapshot = Get-AnimationSnapshot $slide
    $changes += [pscustomobject]@{ kind = 'animation'; target = 'slide:' + $slideIndex; detail = '已配置进入、强调、退出或路径动画及顺序和触发方式' }
  } elseif ($_operation -eq 'configureSlideShow') {
    $settings = $pres.SlideShowSettings
    $settings.ShowType = switch ([string]$actionParams.showType) { 'window' { 2 } 'kiosk' { 3 } default { 1 } }
    $settings.AdvanceMode = if ($actionParams.autoPlay -eq $true -or $actionParams.useSlideTimings -eq $true) { 2 } else { 1 }
    $settings.LoopUntilStopped = if ($actionParams.loop -eq $true) { -1 } else { 0 }
    $settings.ShowWithAnimation = if ($actionParams.showWithAnimation -eq $false) { 0 } else { -1 }
    $targetSlides = if ($actionParams.allSlides -eq $false) { @($slide) } else { @($pres.Slides) }
    $entryEffect = switch ([string]$actionParams.transition) { 'cut' { 257 } 'dissolve' { 1537 } 'wipe' { 2817 } 'none' { 0 } default { 1793 } }
    foreach ($targetSlide in $targetSlides) {
      $transition = $targetSlide.SlideShowTransition
      $transition.EntryEffect = $entryEffect
      $transition.AdvanceOnClick = if ($actionParams.advanceOnClick -eq $false) { 0 } else { -1 }
      $transition.AdvanceOnTime = if ($actionParams.autoPlay -eq $true) { -1 } else { 0 }
      if ($actionParams.advanceAfter) { $transition.AdvanceTime = [double]$actionParams.advanceAfter }
      if ($actionParams.transitionDuration) { try { $transition.Duration = [double]$actionParams.transitionDuration } catch {} }
    }
    $operationData.slideShow = [pscustomobject]@{ showType = [int]$settings.ShowType; advanceMode = [int]$settings.AdvanceMode; loop = [bool]$settings.LoopUntilStopped; slides = @($targetSlides).Count }
    $changes += [pscustomobject]@{ kind = 'slide-show'; target = 'presentation'; detail = '已配置自动播放、循环放映和页面切换' }
  } elseif ($_operation -eq 'setSpeakerNotes') {
    $written = @()
    if ($actionParams.notesBySlide) {
      foreach ($note in @($actionParams.notesBySlide)) {
        if (-not $note.slideIndex -or $note.slideIndex -gt $pres.Slides.Count) { continue }
        $targetSlide = $pres.Slides.Item([int]$note.slideIndex)
        $existingNotes = Get-SlideNotesText $targetSlide
        $applied = Set-SlideNotesText $targetSlide ([string]$note.text) ($note.append -eq $true)
        if ($progId -match '(?i)wpp') {
          $finalNotes = if ($applied) { Get-SlideNotesText $targetSlide } elseif ($note.append -eq $true -and -not [string]::IsNullOrWhiteSpace($existingNotes)) { $existingNotes.Trim() + [Environment]::NewLine + [string]$note.text } else { [string]$note.text }
          $pendingNotesPackageUpdates += [pscustomobject]@{ slideIndex = [int]$note.slideIndex; text = $finalNotes; append = $false }
          $written += [int]$note.slideIndex
        } elseif ($applied) { $written += [int]$note.slideIndex }
        else { $pendingNotesPackageUpdates += [pscustomobject]@{ slideIndex = [int]$note.slideIndex; text = [string]$note.text; append = $note.append -eq $true }; $written += [int]$note.slideIndex }
      }
    } else {
      if ($null -eq $actionParams.text) { throw 'setSpeakerNotes 需要 params.text 或 params.notesBySlide' }
      $existingNotes = Get-SlideNotesText $slide
      $applied = Set-SlideNotesText $slide ([string]$actionParams.text) ($actionParams.append -eq $true)
      if ($progId -match '(?i)wpp') {
        $finalNotes = if ($applied) { Get-SlideNotesText $slide } elseif ($actionParams.append -eq $true -and -not [string]::IsNullOrWhiteSpace($existingNotes)) { $existingNotes.Trim() + [Environment]::NewLine + [string]$actionParams.text } else { [string]$actionParams.text }
        $pendingNotesPackageUpdates += [pscustomobject]@{ slideIndex = $slideIndex; text = $finalNotes; append = $false }
        $written += $slideIndex
      } elseif ($applied) { $written += $slideIndex }
      else { $pendingNotesPackageUpdates += [pscustomobject]@{ slideIndex = $slideIndex; text = [string]$actionParams.text; append = $actionParams.append -eq $true }; $written += $slideIndex }
    }
    $operationData.updatedSlides = $written
    $operationData.notesPackageFallback = $pendingNotesPackageUpdates.Count -gt 0
    $changes += [pscustomobject]@{ kind = 'speaker-notes'; target = 'slides'; detail = '已写入演讲者备注或讲稿' }
  } elseif ($_operation -eq 'inspectSpeakerNotes') {
    $targetSlides = if ($actionParams.allSlides -eq $false) { @($slide) } else { @($pres.Slides) }
    $notes = @()
    foreach ($targetSlide in $targetSlides) {
      $slideText = Get-SlideVisibleText $targetSlide
      $notesText = Get-SlideNotesText $targetSlide
      $notes += [pscustomobject]@{
        slideIndex = [int]$targetSlide.SlideIndex
        slideText = $slideText
        notesText = $notesText
        hasNotes = -not [string]::IsNullOrWhiteSpace($notesText)
        correspondenceScore = Get-NotesCorrespondenceScore $slideText $notesText
      }
    }
    $operationData.notes = $notes
    $operationData.summary = [pscustomobject]@{ slideCount = $notes.Count; missingNotes = @($notes | Where-Object { -not $_.hasNotes }).Count; lowCorrespondence = @($notes | Where-Object { $_.hasNotes -and $_.correspondenceScore -lt 0.2 }).Count }
    $changes = @()
  } else {
    $layout = if ($actionParams.includeNotes -eq $true -or $actionParams.layout -eq 'notes') { 5 } else { switch ([string]$actionParams.layout) { 'one' { 10 } 'two' { 2 } 'three' { 3 } 'four' { 8 } 'six' { 4 } 'nine' { 9 } 'outline' { 6 } default { 3 } } }
    $pres.PrintOptions.OutputType = $layout
    $exported = $false
    try {
      Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint
      Add-Type -AssemblyName office
      if (-not ('WenggePowerPointFixedExporter' -as [type])) {
        $references = @(
          [Microsoft.Office.Interop.PowerPoint._Presentation].Assembly.Location,
          [Microsoft.Office.Core.MsoTriState].Assembly.Location
        )
        $exporterSource = @'
using System;
using System.Runtime.InteropServices;
using P = Microsoft.Office.Interop.PowerPoint;
using O = Microsoft.Office.Core;
public static class WenggePowerPointFixedExporter {
  public static void Export(object value, string path, int outputType) {
    IntPtr pointer = Marshal.GetIUnknownForObject(value);
    try {
      P._Presentation presentation = (P._Presentation)Marshal.GetTypedObjectForIUnknown(pointer, typeof(P._Presentation));
      presentation.ExportAsFixedFormat(
        path,
        P.PpFixedFormatType.ppFixedFormatTypePDF,
        P.PpFixedFormatIntent.ppFixedFormatIntentPrint,
        O.MsoTriState.msoFalse,
        P.PpPrintHandoutOrder.ppPrintHandoutVerticalFirst,
        (P.PpPrintOutputType)outputType,
        O.MsoTriState.msoFalse,
        null,
        P.PpPrintRangeType.ppPrintAll,
        null,
        true, true, true, true, false,
        Type.Missing
      );
    } finally {
      Marshal.Release(pointer);
    }
  }
}
'@
        Add-Type -ReferencedAssemblies $references -TypeDefinition $exporterSource
      }
      [WenggePowerPointFixedExporter]::Export($pres, $_outputPath, $layout)
      $exported = $true
    } catch {
      try {
        $missing = [Type]::Missing
        $pres.ExportAsFixedFormat($_outputPath, 2, 2, 0, 1, $layout, 0, $missing, 1, '', $true, $true, $true, $true, $false, $missing)
        $exported = $true
      } catch {
        throw '当前 PowerPoint/WPS 版本无法导出指定的备注或讲义版式: ' + $_.Exception.Message
      }
    }
    if (-not $exported -or -not [IO.File]::Exists($_outputPath)) { throw 'PowerPoint 讲义 PDF 未生成' }
    $operationData.outputType = $layout
    $operationData.includeNotes = $layout -eq 5
    $changes += [pscustomobject]@{ kind = 'export'; target = $_outputPath; detail = '已导出带备注或讲义版 PDF' }
  }
`;
}
