using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using D = DocumentFormat.OpenXml.Drawing;
using S = DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.OpenXml;

internal sealed class OpenXmlTableService
{
    public object Inspect(string filePath, string? target)
    {
        var path = RequireFile(filePath);
        var extension = Path.GetExtension(path).ToLowerInvariant();
        var tables = extension switch
        {
            ".docx" or ".docm" => InspectWord(path),
            ".xlsx" or ".xlsm" => InspectSpreadsheet(path),
            ".pptx" or ".pptm" => InspectPresentation(path),
            _ => throw Unsupported(extension),
        };
        return new
        {
            engine = "openxml",
            operation = "inspectTable",
            documentType = TypeName(extension),
            filePath = path,
            target,
            tableCount = tables.Count,
            tables,
        };
    }

    public object ApplyStyle(string filePath, string style, string? outputPath, string? target)
    {
        var source = RequireFile(filePath);
        var output = string.IsNullOrWhiteSpace(outputPath)
            ? Path.Combine(Path.GetDirectoryName(source)!, $"{Path.GetFileNameWithoutExtension(source)}-styled{Path.GetExtension(source)}")
            : Path.GetFullPath(outputPath);
        Directory.CreateDirectory(Path.GetDirectoryName(output) ?? Environment.CurrentDirectory);
        if (!SamePath(source, output)) File.Copy(source, output, overwrite: true);
        var color = style switch { "financial" => "385723", "compact" => "5B9BD5", _ => "1F4E79" };
        var extension = Path.GetExtension(output).ToLowerInvariant();
        var changedParts = extension switch
        {
            ".docx" or ".docm" => StyleWord(output, color),
            ".xlsx" or ".xlsm" => StyleSpreadsheet(output, color),
            ".pptx" or ".pptm" => StylePresentation(output, color),
            _ => throw Unsupported(extension),
        };
        return new
        {
            engine = "openxml",
            operation = "applyTableStyle",
            documentType = TypeName(extension),
            filePath = source,
            outputPath = output,
            target,
            style,
            changedParts,
        };
    }

    private static List<object> InspectWord(string path)
    {
        using var document = WordprocessingDocument.Open(path, false);
        return document.MainDocumentPart?.Document.Descendants<W.Table>().Select((table, index) =>
        {
            var rows = table.Elements<W.TableRow>().Select((row, rowIndex) =>
            {
                var cells = row.Elements<W.TableCell>().Select((cell, columnIndex) => new
                {
                    text = cell.InnerText,
                    rowIndex,
                    columnIndex,
                    bold = cell.Descendants<W.Bold>().Any(),
                    fillColor = cell.TableCellProperties?.Shading?.Fill?.Value,
                }).ToArray();
                return new { rowIndex, isHeaderGuess = rowIndex == 0 && HeaderGuess(cells.Select(cell => cell.text)), cells };
            }).ToArray();
            return (object)new { index, partName = "word/document.xml", rows, columns = rows.Select(row => row.cells.Length).DefaultIfEmpty().Max() };
        }).ToList() ?? [];
    }

    private static List<object> InspectPresentation(string path)
    {
        using var document = PresentationDocument.Open(path, false);
        var tables = new List<object>();
        foreach (var slide in document.PresentationPart?.SlideParts ?? [])
        {
            foreach (var table in slide.Slide.Descendants<D.Table>())
            {
                var rows = table.Elements<D.TableRow>().Select((row, rowIndex) =>
                {
                    var cells = row.Elements<D.TableCell>().Select((cell, columnIndex) => new
                    {
                        text = cell.InnerText,
                        rowIndex,
                        columnIndex,
                        bold = cell.Descendants<D.RunProperties>().Any(properties => properties.Bold?.Value == true),
                    }).ToArray();
                    return new { rowIndex, isHeaderGuess = rowIndex == 0 && HeaderGuess(cells.Select(cell => cell.text)), cells };
                }).ToArray();
                tables.Add(new { index = tables.Count, partName = slide.Uri.ToString().TrimStart('/'), rows, columns = rows.Select(row => row.cells.Length).DefaultIfEmpty().Max() });
            }
        }
        return tables;
    }

    private static List<object> InspectSpreadsheet(string path)
    {
        using var document = SpreadsheetDocument.Open(path, false);
        var tables = new List<object>();
        foreach (var worksheet in document.WorkbookPart?.WorksheetParts ?? [])
        {
            var rows = worksheet.Worksheet.Descendants<S.Row>().Select((row, rowIndex) =>
            {
                var cells = row.Elements<S.Cell>().Select((cell, columnIndex) => new
                {
                    text = CellText(document, cell),
                    rowIndex,
                    columnIndex,
                    reference = cell.CellReference?.Value,
                }).ToArray();
                return new { rowIndex, isHeaderGuess = rowIndex == 0 && HeaderGuess(cells.Select(cell => cell.text)), cells };
            }).ToArray();
            if (rows.Length > 0)
                tables.Add(new { index = tables.Count, partName = worksheet.Uri.ToString().TrimStart('/'), rows, columns = rows.Select(row => row.cells.Length).DefaultIfEmpty().Max() });
        }
        return tables;
    }

    private static List<string> StyleWord(string path, string color)
    {
        using var document = WordprocessingDocument.Open(path, true);
        var changed = false;
        foreach (var table in document.MainDocumentPart?.Document.Descendants<W.Table>() ?? [])
        {
            var header = table.Elements<W.TableRow>().FirstOrDefault();
            if (header is null) continue;
            foreach (var cell in header.Elements<W.TableCell>())
            {
                cell.TableCellProperties ??= new W.TableCellProperties();
                cell.TableCellProperties.Shading = new W.Shading { Fill = color };
                foreach (var run in cell.Descendants<W.Run>())
                {
                    run.RunProperties ??= new W.RunProperties();
                    run.RunProperties.Bold = new W.Bold();
                }
            }
            changed = true;
        }
        if (changed) document.MainDocumentPart!.Document.Save();
        return changed ? ["word/document.xml"] : [];
    }

    private static List<string> StylePresentation(string path, string color)
    {
        using var document = PresentationDocument.Open(path, true);
        var changed = new List<string>();
        foreach (var slide in document.PresentationPart?.SlideParts ?? [])
        {
            var slideChanged = false;
            foreach (var table in slide.Slide.Descendants<D.Table>())
            {
                var header = table.Elements<D.TableRow>().FirstOrDefault();
                if (header is null) continue;
                foreach (var cell in header.Elements<D.TableCell>())
                {
                    cell.TableCellProperties ??= new D.TableCellProperties();
                    cell.TableCellProperties.RemoveAllChildren<D.SolidFill>();
                    cell.TableCellProperties.PrependChild(new D.SolidFill(new D.RgbColorModelHex { Val = color }));
                    foreach (var run in cell.Descendants<D.Run>())
                    {
                        run.RunProperties ??= new D.RunProperties();
                        run.RunProperties.Bold = true;
                    }
                }
                slideChanged = true;
            }
            if (!slideChanged) continue;
            slide.Slide.Save();
            changed.Add(slide.Uri.ToString().TrimStart('/'));
        }
        return changed;
    }

    private static List<string> StyleSpreadsheet(string path, string color)
    {
        using var document = SpreadsheetDocument.Open(path, true);
        var workbook = document.WorkbookPart ?? throw new OfficeWorkerException("openxml_invalid", "工作簿缺少 WorkbookPart");
        var styles = workbook.WorkbookStylesPart ?? workbook.AddNewPart<WorkbookStylesPart>();
        styles.Stylesheet ??= CreateStylesheet();
        var styleIndex = AppendHeaderStyle(styles.Stylesheet, color);
        styles.Stylesheet.Save();
        var changed = new List<string> { "xl/styles.xml" };
        foreach (var worksheet in workbook.WorksheetParts)
        {
            var row = worksheet.Worksheet.Descendants<S.Row>().FirstOrDefault();
            if (row is null) continue;
            foreach (var cell in row.Elements<S.Cell>()) cell.StyleIndex = styleIndex;
            worksheet.Worksheet.Save();
            changed.Add(worksheet.Uri.ToString().TrimStart('/'));
        }
        return changed;
    }

    private static S.Stylesheet CreateStylesheet() => new(
        new S.Fonts(new S.Font()) { Count = 1 },
        new S.Fills(new S.Fill(new S.PatternFill { PatternType = S.PatternValues.None }), new S.Fill(new S.PatternFill { PatternType = S.PatternValues.Gray125 })) { Count = 2 },
        new S.Borders(new S.Border()) { Count = 1 },
        new S.CellStyleFormats(new S.CellFormat()) { Count = 1 },
        new S.CellFormats(new S.CellFormat()) { Count = 1 },
        new S.CellStyles(new S.CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 }) { Count = 1 });

    private static uint AppendHeaderStyle(S.Stylesheet stylesheet, string color)
    {
        stylesheet.Fonts ??= new S.Fonts();
        stylesheet.Fills ??= new S.Fills();
        stylesheet.Borders ??= new S.Borders(new S.Border());
        stylesheet.CellFormats ??= new S.CellFormats(new S.CellFormat());
        var fontId = (uint)stylesheet.Fonts.ChildElements.Count;
        stylesheet.Fonts.Append(new S.Font(new S.Bold()));
        stylesheet.Fonts.Count = (uint)stylesheet.Fonts.ChildElements.Count;
        var fillId = (uint)stylesheet.Fills.ChildElements.Count;
        stylesheet.Fills.Append(new S.Fill(new S.PatternFill(new S.ForegroundColor { Rgb = $"FF{color}" }, new S.BackgroundColor { Indexed = 64 }) { PatternType = S.PatternValues.Solid }));
        stylesheet.Fills.Count = (uint)stylesheet.Fills.ChildElements.Count;
        var index = (uint)stylesheet.CellFormats.ChildElements.Count;
        stylesheet.CellFormats.Append(new S.CellFormat { FontId = fontId, FillId = fillId, BorderId = 0, FormatId = 0, ApplyFont = true, ApplyFill = true });
        stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.ChildElements.Count;
        return index;
    }

    private static string CellText(SpreadsheetDocument document, S.Cell cell)
    {
        if (cell.DataType?.Value == S.CellValues.SharedString && int.TryParse(cell.CellValue?.Text, out var index))
            return document.WorkbookPart?.SharedStringTablePart?.SharedStringTable?.Elements<S.SharedStringItem>().ElementAtOrDefault(index)?.InnerText ?? string.Empty;
        return cell.InlineString?.InnerText ?? cell.CellValue?.Text ?? string.Empty;
    }

    private static bool HeaderGuess(IEnumerable<string> values)
    {
        var cells = values.Where(value => !string.IsNullOrWhiteSpace(value)).ToArray();
        return cells.Length > 0 && cells.Count(value => !double.TryParse(value, out _)) >= Math.Ceiling(cells.Length / 2d);
    }

    private static string RequireFile(string path) => File.Exists(Path.GetFullPath(path)) ? Path.GetFullPath(path) : throw new OfficeWorkerException("file_not_found", $"Office 文件不存在: {Path.GetFullPath(path)}");
    private static bool SamePath(string left, string right) => string.Equals(Path.GetFullPath(left), Path.GetFullPath(right), StringComparison.OrdinalIgnoreCase);
    private static string TypeName(string extension) => extension is ".docx" or ".docm" ? "word" : extension is ".xlsx" or ".xlsm" ? "spreadsheet" : "presentation";
    private static OfficeWorkerException Unsupported(string extension) => new("unsupported_file", $"不支持的 Open XML 文件类型: {extension}");
}
