/**
 * Office COM 高级 action 兜底执行器。
 *
 * 关联模块：
 * - officeCore/officeActionAdapter.ts: Open XML 不支持或显式 preferEngine=com 时转到这里。
 * - automation/powershell.ts: 负责 PowerShell 执行和变量安全注入。
 */

import path from "path";
import { executePowerShell, psVar } from "../../../automation/powershell";
import { safeJsonParse } from "../../../automation/json";
import { parseOfficeLocator } from "../../officeCore/locator";
import { doneResult, failedResult } from "../../officeCore/results";
import type { OfficeActionBridge } from "../../contracts/office";
import type { OfficeActionInput, OfficeActionResult } from "../../officeCore/types";

type ComChange = { kind: string; target?: string; detail: string };

const EXCEL_COM_OPERATIONS = new Set(["insertChart", "applyConditionalFormatting", "setDataValidation", "styleTable"]);
const WORD_COM_OPERATIONS = new Set(["applyHeadingStyles", "insertOrUpdateToc", "styleTables", "setHeaderFooter", "insertOrReplaceImage", "snapshot"]);
const PPT_COM_OPERATIONS = new Set(["applyTheme", "deleteSlides", "normalizeLayouts", "insertChart", "replacePictureSlot", "alignShapes", "snapshot"]);

export class OfficeComActionBridge implements OfficeActionBridge {
  async executeAction(input: OfficeActionInput): Promise<OfficeActionResult> {
    if (!input.filePath) {
      return failedResult({ ...input, preferEngine: "com" }, "缺少 filePath，无法执行 COM Office action");
    }

    if (!supportsComAction(input)) {
      return {
        status: "unsupported",
        engine: "com",
        app: input.app,
        action: input.action,
        operation: input.operation,
        filePath: input.filePath,
        outputPath: input.outputPath,
        target: input.target,
        summary: `暂不支持 COM Office action: ${input.app}/${input.operation}`,
        changes: [],
      };
    }

    try {
      const script = buildComScript(input);
      const output = await executePowerShell(script, 120000);
      const data = safeJsonParse<{ outputPath?: string; changes?: ComChange[] }>(output, "powershell", "执行 COM Office action");
      const outputPath = data.outputPath || input.outputPath || input.filePath;
      return doneResult({
        engine: "com",
        app: input.app,
        action: input.action,
        operation: input.operation,
        filePath: input.filePath,
        outputPath,
        target: input.target,
        summary: `已通过 COM 执行 ${input.app === "presentation" ? "PowerPoint" : input.app[0].toUpperCase() + input.app.slice(1)} ${input.operation}`,
        data,
        validation: {
          ok: true,
          checks: [{ name: "com-route", ok: true, message: "COM 兜底执行完成" }],
        },
        changes: Array.isArray(data.changes) ? data.changes : [],
      });
    } catch (error) {
      return failedResult({ ...input, preferEngine: "com" }, error);
    }
  }
}

function supportsComAction(input: OfficeActionInput): boolean {
  if (input.app === "excel") return EXCEL_COM_OPERATIONS.has(input.operation);
  if (input.app === "word") return WORD_COM_OPERATIONS.has(input.operation);
  return PPT_COM_OPERATIONS.has(input.operation);
}

function buildComScript(input: OfficeActionInput): string {
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

function stringParam(input: OfficeActionInput, key: string): string | undefined {
  return typeof input.params?.[key] === "string" ? input.params[key] : undefined;
}

function validationFormulaParam(input: OfficeActionInput): string {
  const values = input.params?.values;
  if (Array.isArray(values)) return values.map((item) => String(item)).join(",");
  return stringParam(input, "formula") || stringParam(input, "values") || "";
}

function targetSlideIndex(target?: string): number {
  const locator = target ? parseOfficeLocator(target) : undefined;
  return locator?.kind === "slide" && locator.index ? locator.index : 1;
}

function presentationDeleteSlideIndexesLiteral(input: OfficeActionInput): string {
  const slides = Array.isArray(input.params?.slides)
    ? input.params.slides.map((item) => Number(item)).filter((item) => Number.isFinite(item))
    : [];
  const range = slides.length > 0 ? slides : presentationDeleteRange(input);
  const indexes = [...new Set(range.map((item) => Math.floor(item)).filter((item) => item >= 1))].sort((a, b) => a - b);
  return indexes.length > 0 ? `@(${indexes.join(", ")})` : "@()";
}

function presentationDeleteRange(input: OfficeActionInput): number[] {
  const from = numericParam(input, "from") ?? numericParam(input, "start");
  const to = numericParam(input, "to") ?? numericParam(input, "end") ?? from;
  if (from && to) return buildNumberRange(from, to);
  const target = input.target || "";
  const match = target.match(/^slides?:\s*(\d+)(?:\s*-\s*(\d+))?$/i);
  if (!match) return [];
  return buildNumberRange(Number(match[1]), Number(match[2] || match[1]));
}

function buildNumberRange(from: number, to: number): number[] {
  const start = Math.floor(Math.min(from, to));
  const end = Math.floor(Math.max(from, to));
  return Array.from({ length: end - start + 1 }, (_, index) => start + index);
}

function numericParam(input: OfficeActionInput, key: string): number | undefined {
  const value = input.params?.[key];
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

function normalizeHeadingLevel(value: unknown): number {
  return typeof value === "number" && value >= 1 && value <= 9 ? Math.floor(value) : 1;
}

function excelChartType(value?: string): number {
  switch ((value || "column").toLowerCase()) {
    case "line": return 4;
    case "pie": return 5;
    case "bar": return 57;
    case "area": return 1;
    case "scatter": return -4169;
    default: return 51;
  }
}

function pptChartType(value?: string): number {
  return excelChartType(value);
}

function colorToOle(hex: string): number {
  const normalized = /^[0-9a-fA-F]{6}$/.test(hex) ? hex : "1F4E79";
  const r = parseInt(normalized.slice(0, 2), 16);
  const g = parseInt(normalized.slice(2, 4), 16);
  const b = parseInt(normalized.slice(4, 6), 16);
  return r + (g * 256) + (b * 65536);
}

function defaultOutputPath(filePath: string, suffix: string): string {
  if (suffix.includes(".")) {
    return path.join(path.dirname(filePath), `${path.basename(filePath, path.extname(filePath))}-${suffix}`);
  }
  const ext = path.extname(filePath);
  return path.join(path.dirname(filePath), `${path.basename(filePath, ext)}-${suffix}${ext}`);
}

function psLiteralText(value: string): string {
  return value.replace(/'/g, "''");
}

function excelChangeKind(operation: string): string {
  if (operation === "insertChart") return "chart";
  if (operation === "setDataValidation") return "validation";
  if (operation === "applyConditionalFormatting") return "conditional-format";
  return "table-style";
}

function excelChangeDetail(operation: string): string {
  if (operation === "insertChart") return "已创建图表";
  if (operation === "setDataValidation") return "已设置数据验证";
  if (operation === "applyConditionalFormatting") return "已应用条件格式";
  return "已应用表格样式";
}

function wordChangeKind(operation: string): string {
  if (operation === "insertOrUpdateToc") return "toc";
  if (operation === "insertOrReplaceImage") return "image";
  if (operation === "snapshot") return "snapshot";
  return "document-style";
}

function wordChangeDetail(operation: string): string {
  if (operation === "insertOrUpdateToc") return "已插入或更新目录";
  if (operation === "insertOrReplaceImage") return "已插入图片";
  if (operation === "snapshot") return "已导出文档预览";
  return "已应用 Word 样式";
}

function presentationChangeKind(operation: string): string {
  if (operation === "snapshot") return "snapshot";
  if (operation === "deleteSlides") return "slide-delete";
  if (operation === "insertChart") return "chart";
  if (operation === "replacePictureSlot") return "image";
  return "presentation-style";
}

function presentationChangeDetail(operation: string): string {
  if (operation === "snapshot") return "已导出幻灯片快照";
  if (operation === "deleteSlides") return "已删除幻灯片";
  if (operation === "insertChart") return "已插入图表";
  if (operation === "replacePictureSlot") return "已替换图片占位";
  if (operation === "alignShapes") return "已对齐形状";
  if (operation === "normalizeLayouts") return "已规范版式";
  return "已应用 PPT 主题";
}
