import { psVar } from "../../../automation/powershell";
import { parseOfficeLocator } from "../../officeCore/locator";
import type { OfficeActionInput } from "../../officeCore/types";
import {
  colorToOle,
  defaultOutputPath,
  excelChangeDetail,
  excelChangeKind,
  excelChartType,
  normalizeHeadingLevel,
  pptChartType,
  presentationChangeDetail,
  presentationChangeKind,
  presentationDeleteSlideIndexesLiteral,
  psLiteralText,
  stringParam,
  targetSlideIndex,
  validationFormulaParam,
  wordChangeDetail,
  wordChangeKind,
} from "./officeComActionScriptHelpers";

export function buildComScript(input: OfficeActionInput): string {
  if (input.app === "excel") return buildExcelScript(input);
  if (input.app === "word") return buildWordScript(input);
  return buildPresentationScript(input);
}

function buildExcelScript(input: OfficeActionInput): string {
  const locator = parseOfficeLocator(input.target || "");
  const sheetName = locator.sheetName || stringParam(input, "sheetName") || "Sheet1";
  const rangeAddress = locator.address || stringParam(input, "range") || "A1";
  const chartType = excelChartType(stringParam(input, "chartType"));
  const fillColor = colorToOle(stringParam(input, "fillColor") || "FFF2CC");
  const styleColor = colorToOle(stringParam(input, "headerColor") || "1F4E79");
  const validationFormula = validationFormulaParam(input);
  const cfFormula = stringParam(input, "formula") || "TRUE";
  const outputPath = input.outputPath || defaultOutputPath(input.filePath!, "com");

  return `
${psVar("_filePath", input.filePath!)}
${psVar("_outputPath", outputPath)}
${psVar("_sheetName", sheetName)}
${psVar("_rangeAddress", rangeAddress)}
${psVar("_validationFormula", validationFormula)}
${psVar("_cfFormula", cfFormula)}
$chartType = ${chartType}
$fillColor = ${fillColor}
$styleColor = ${styleColor}
$app = $null
$workbook = $null
try {
  $app = New-Object -ComObject Excel.Application
  $app.Visible = $true
  $workbook = $app.Workbooks.Open($_filePath)
  $sheet = $workbook.Worksheets.Item($_sheetName)
  $range = $sheet.Range($_rangeAddress)
${excelOperationScript(input.operation)}
  if ($_outputPath -and $_outputPath -ne $_filePath) {
    $workbook.SaveAs($_outputPath)
  } else {
    $workbook.Save()
    $_outputPath = $_filePath
  }
  [pscustomobject]@{
    outputPath = $_outputPath
    changes = @([pscustomobject]@{ kind = '${excelChangeKind(input.operation)}'; target = '${psLiteralText(input.target || rangeAddress)}'; detail = '${excelChangeDetail(input.operation)}' })
  } | ConvertTo-Json -Depth 6 -Compress
} finally {
  if ($null -ne $workbook) {
    try { $workbook.Close($true) } catch {}
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) } catch {}
  }
  if ($null -ne $app) {
    try { $app.Quit() } catch {}
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch {}
  }
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}
`;
}

function excelOperationScript(operation: string): string {
  switch (operation) {
    case "insertChart":
      return `
  $chartObjects = $sheet.ChartObjects()
  $chartObject = $chartObjects.Add($range.Left + $range.Width + 20, $range.Top, 420, 260)
  $chart = $chartObject.Chart
  $chart.SetSourceData($range)
  $chart.ChartType = $chartType
`;
    case "setDataValidation":
      return `
  try { $range.Validation.Delete() } catch {}
  $range.Validation.Add(3, 1, 1, $_validationFormula)
  $range.Validation.IgnoreBlank = $true
  $range.Validation.InCellDropdown = $true
`;
    case "applyConditionalFormatting":
      return `
  $condition = $range.FormatConditions.Add(2, 3, $_cfFormula)
  $condition.Interior.Color = $fillColor
`;
    default:
      return `
  $range.Font.Bold = $true
  $range.Borders.LineStyle = 1
  $range.Interior.Color = $styleColor
`;
  }
}

function buildWordScript(input: OfficeActionInput): string {
  const outputPath = input.outputPath || defaultOutputPath(input.filePath!, "com");
  const startsWith = stringParam(input, "startsWith") || "";
  const headerFooterText = stringParam(input, "text") || "";
  const headerFooterKind = stringParam(input, "kind") === "footer" ? "footer" : "header";
  const imagePath = stringParam(input, "imagePath") || "";
  const headingLevel = normalizeHeadingLevel(input.params?.level);

  return `
${psVar("_filePath", input.filePath!)}
${psVar("_outputPath", outputPath)}
${psVar("_startsWith", startsWith)}
${psVar("_headerFooterText", headerFooterText)}
${psVar("_headerFooterKind", headerFooterKind)}
${psVar("_imagePath", imagePath)}
$headingLevel = ${headingLevel}
$app = $null
$doc = $null
try {
  $app = New-Object -ComObject Word.Application
  $app.Visible = $true
  $doc = $app.Documents.Open($_filePath)
${wordOperationScript(input.operation)}
  if ($_outputPath -and $_outputPath -ne $_filePath) {
    if ('${input.operation}' -eq 'snapshot') {
      $doc.ExportAsFixedFormat($_outputPath, 17)
    } else {
      $doc.SaveAs2($_outputPath)
    }
  } else {
    $doc.Save()
    $_outputPath = $_filePath
  }
  [pscustomobject]@{
    outputPath = $_outputPath
    changes = @([pscustomobject]@{ kind = '${wordChangeKind(input.operation)}'; target = '${psLiteralText(input.target || "")}'; detail = '${wordChangeDetail(input.operation)}' })
  } | ConvertTo-Json -Depth 6 -Compress
} finally {
  if ($null -ne $doc) {
    try { $doc.Close($false) } catch {}
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) } catch {}
  }
  if ($null -ne $app) {
    try { $app.Quit() } catch {}
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch {}
  }
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}
`;
}

function wordOperationScript(operation: string): string {
  switch (operation) {
    case "applyHeadingStyles":
      return `
  foreach ($paragraph in $doc.Paragraphs) {
    $text = ([string]$paragraph.Range.Text).Trim()
    if (-not $_startsWith -or $text.StartsWith($_startsWith)) {
      $styleId = -1 - $headingLevel
      $paragraph.Range.Style = $doc.Styles.Item($styleId)
    }
  }
`;
    case "insertOrUpdateToc":
      return `
  if ($doc.TablesOfContents.Count -gt 0) {
    foreach ($toc in $doc.TablesOfContents) { $toc.Update() }
  } else {
    $range = $doc.Range(0, 0)
    $toc = $doc.TablesOfContents.Add($range, $true, 1, 3)
    $toc.Update()
  }
  $doc.Fields.Update()
`;
    case "styleTables":
      return `
  foreach ($table in $doc.Tables) {
    $table.Borders.Enable = 1
    if ($table.Rows.Count -gt 0) {
      $table.Rows.Item(1).Range.Bold = $true
      $table.Rows.Item(1).Shading.BackgroundPatternColor = 15773696
    }
  }
`;
    case "setHeaderFooter":
      return `
  foreach ($section in $doc.Sections) {
    if ($_headerFooterKind -eq 'footer') {
      $section.Footers.Item(1).Range.Text = $_headerFooterText
    } else {
      $section.Headers.Item(1).Range.Text = $_headerFooterText
    }
  }
`;
    case "insertOrReplaceImage":
      return `
  if ([string]::IsNullOrWhiteSpace($_imagePath)) { throw 'insertOrReplaceImage 需要 params.imagePath' }
  $range = $doc.Range([Math]::Max(0, $doc.Content.End - 1), [Math]::Max(0, $doc.Content.End - 1))
  [void]$doc.InlineShapes.AddPicture($_imagePath, $false, $true, $range)
`;
    default:
      return "  $doc.Fields.Update()";
  }
}

function buildPresentationScript(input: OfficeActionInput): string {
  const outputPath = input.outputPath || defaultOutputPath(input.filePath!, input.operation === "snapshot" ? "snapshot.png" : "com");
  const slideIndex = targetSlideIndex(input.target);
  const accentColor = colorToOle(stringParam(input, "accentColor") || "1F4E79");
  const chartType = pptChartType(stringParam(input, "chartType"));
  const imagePath = stringParam(input, "imagePath") || "";
  const deleteSlideIndexes = presentationDeleteSlideIndexesLiteral(input);

  return `
${psVar("_filePath", input.filePath!)}
${psVar("_outputPath", outputPath)}
${psVar("_imagePath", imagePath)}
$slideIndex = ${slideIndex}
$deleteSlideIndexes = ${deleteSlideIndexes}
$accentColor = ${accentColor}
$chartType = ${chartType}
$app = $null
$pres = $null
try {
  $app = New-Object -ComObject PowerPoint.Application
  $app.Visible = $true
  $pres = $app.Presentations.Open($_filePath)
  if ($slideIndex -gt $pres.Slides.Count) { throw "幻灯片序号超出范围: $slideIndex" }
  $slide = $pres.Slides.Item($slideIndex)
${presentationOperationScript(input.operation)}
  if ('${input.operation}' -ne 'snapshot') {
    if ($_outputPath -and $_outputPath -ne $_filePath) {
      $pres.SaveAs($_outputPath)
    } else {
      $pres.Save()
      $_outputPath = $_filePath
    }
  }
  [pscustomobject]@{
    outputPath = $_outputPath
    changes = @([pscustomobject]@{ kind = '${presentationChangeKind(input.operation)}'; target = 'slide:' + $slideIndex; detail = '${presentationChangeDetail(input.operation)}' })
  } | ConvertTo-Json -Depth 6 -Compress
} finally {
  if ($null -ne $pres) {
    try { $pres.Close() } catch {}
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) } catch {}
  }
  if ($null -ne $app) {
    try { $app.Quit() } catch {}
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch {}
  }
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}
`;
}

function presentationOperationScript(operation: string): string {
  switch (operation) {
    case "snapshot":
      return "  $slide.Export($_outputPath, 'PNG')";
    case "deleteSlides":
      return `
  if ($deleteSlideIndexes.Count -eq 0) { throw 'deleteSlides 需要 params.slides、params.from/to 或 target: slide:2-6' }
  foreach ($idx in ($deleteSlideIndexes | Sort-Object -Descending)) {
    if ($idx -lt 1 -or $idx -gt $pres.Slides.Count) { throw "幻灯片序号超出范围: $idx" }
    if ($pres.Slides.Count -le 1) { throw 'deleteSlides 至少需要保留一张幻灯片' }
    $pres.Slides.Item($idx).Delete()
  }
`;
    case "insertChart":
      return `
  $shape = $slide.Shapes.AddChart2(201, $chartType, 80, 120, 520, 300)
  $shape.Name = 'AI Chart'
`;
    case "replacePictureSlot":
      return `
  if ([string]::IsNullOrWhiteSpace($_imagePath)) { throw 'replacePictureSlot 需要 params.imagePath' }
  $slide.Shapes.AddPicture($_imagePath, $false, $true, 80, 120, 520, 300)
`;
    case "alignShapes":
      return `
  foreach ($shape in $slide.Shapes) {
    try {
      if ($shape.Left -lt 80) { $shape.Left = 80 }
      if ($shape.Top -lt 80) { $shape.Top = 80 }
    } catch {}
  }
`;
    case "normalizeLayouts":
      return `
  foreach ($shape in $slide.Shapes) {
    try {
      if ($shape.Width -gt 600) { $shape.Width = 600 }
      if ($shape.Height -gt 360) { $shape.Height = 360 }
    } catch {}
  }
`;
    default:
      return `
  foreach ($shape in $slide.Shapes) {
    try {
      if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
        $shape.TextFrame.TextRange.Font.Color.RGB = $accentColor
      }
    } catch {}
  }
`;
  }
}
