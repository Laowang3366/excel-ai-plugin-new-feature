using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.OpenXml;

internal sealed class OpenXmlDocumentParserService
{
    private const long MaxSpreadsheetBytes = 25 * 1024 * 1024;
    private const int MaxSpreadsheetRows = 500;

    public object Parse(string filePath)
    {
        var fullPath = Path.GetFullPath(filePath);
        if (!File.Exists(fullPath))
        {
            throw new OfficeWorkerException("file_not_found", $"Office 文件不存在: {fullPath}");
        }

        var extension = Path.GetExtension(fullPath).ToLowerInvariant();
        var chunks = extension switch
        {
            ".xlsx" or ".xlsm" => ParseSpreadsheet(fullPath, extension[1..]),
            ".docx" or ".docm" => ParseWord(fullPath, extension[1..]),
            ".pptx" or ".pptm" => ParsePresentation(fullPath, extension[1..]),
            _ => throw new OfficeWorkerException("unsupported_file", $"不支持的 Open XML 文件类型: {extension}"),
        };

        return new { filePath = fullPath, chunks };
    }

    private static IReadOnlyList<ParsedChunk> ParseSpreadsheet(string path, string sourceType)
    {
        var info = new FileInfo(path);
        if (info.Length > MaxSpreadsheetBytes)
        {
            throw new OfficeWorkerException("file_too_large", $"Excel 文件过大，知识库索引最多支持 {MaxSpreadsheetBytes / 1024 / 1024}MB: {info.Name}");
        }

        using var document = SpreadsheetDocument.Open(path, false);
        var workbookPart = document.WorkbookPart;
        if (workbookPart?.Workbook.Sheets is null) return [];
        var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable?
            .Elements<S.SharedStringItem>().Select(item => item.InnerText).ToArray() ?? [];
        var chunks = new List<ParsedChunk>();

        foreach (var sheet in workbookPart.Workbook.Sheets.Elements<S.Sheet>())
        {
            if (sheet.Id?.Value is not string relationshipId || workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart)
            {
                continue;
            }

            var allRows = worksheetPart.Worksheet.GetFirstChild<S.SheetData>()?.Elements<S.Row>().ToArray() ?? [];
            if (allRows.Length == 0) continue;
            var parsedRows = allRows.Take(MaxSpreadsheetRows + 1)
                .Select(row => ReadSpreadsheetRow(row, sharedStrings)).ToArray();
            var headers = parsedRows[0].Select(value => value.Trim()).ToArray();
            var rowCount = Math.Max(0, allRows.Length - 1);
            var maxColumn = allRows.SelectMany(row => row.Elements<S.Cell>())
                .Select(cell => ColumnNumber(cell.CellReference?.Value)).DefaultIfEmpty(0).Max();
            var lines = new List<string> { $"【表头】{string.Join(" | ", headers)}" };
            lines.AddRange(parsedRows.Skip(1).Select(row => string.Join(" | ", row.Select(value => value.Trim()))));
            if (rowCount > MaxSpreadsheetRows)
            {
                lines.Add($"...（还有 {rowCount - MaxSpreadsheetRows} 行未展示）");
            }

            var dimension = worksheetPart.Worksheet.SheetDimension?.Reference?.Value;
            var maxRow = allRows.Select(row => (int)(row.RowIndex?.Value ?? 0)).DefaultIfEmpty(0).Max();
            var range = !string.IsNullOrWhiteSpace(dimension)
                ? dimension
                : $"A1:{CellReference(Math.Max(1, maxColumn), Math.Max(1, maxRow))}";
            chunks.Add(new ParsedChunk(
                string.Join('\n', lines),
                sourceType,
                new ParsedMetadata(
                    SheetName: sheet.Name?.Value ?? $"Sheet{chunks.Count + 1}",
                    TableRange: range,
                    Headers: headers,
                    RowCount: rowCount,
                    ColCount: Math.Max(headers.Length, maxColumn))));
        }

        return chunks;
    }

    private static IReadOnlyList<ParsedChunk> ParseWord(string path, string sourceType)
    {
        using var document = WordprocessingDocument.Open(path, false);
        var main = document.MainDocumentPart;
        if (main is null) return [];
        var roots = new List<OpenXmlElement>();
        if (main.Document.Body is not null) roots.Add(main.Document.Body);
        roots.AddRange(main.HeaderParts.Select(part => part.Header));
        roots.AddRange(main.FooterParts.Select(part => part.Footer));
        if (main.FootnotesPart?.Footnotes is not null) roots.Add(main.FootnotesPart.Footnotes);
        if (main.EndnotesPart?.Endnotes is not null) roots.Add(main.EndnotesPart.Endnotes);

        var lines = roots.SelectMany(root => root.Descendants<W.Paragraph>())
            .Select(paragraph => paragraph.InnerText.Trim()).Where(text => text.Length > 0).ToArray();
        if (lines.Length == 0) return [];
        var rows = roots.SelectMany(root => root.Descendants<W.TableRow>())
            .Select(row => row.Elements<W.TableCell>().Select(cell => cell.InnerText.Trim()).ToArray())
            .Where(row => row.Any(cell => cell.Length > 0)).ToArray();
        return [new ParsedChunk(string.Join('\n', lines), sourceType, new ParsedMetadata(RowCount: lines.Length, Rows: rows))];
    }

    private static IReadOnlyList<ParsedChunk> ParsePresentation(string path, string sourceType)
    {
        using var document = PresentationDocument.Open(path, false);
        var presentationPart = document.PresentationPart;
        var slideIds = presentationPart?.Presentation.SlideIdList?.Elements<P.SlideId>().ToArray() ?? [];
        var chunks = new List<ParsedChunk>();
        for (var index = 0; index < slideIds.Length; index++)
        {
            var relationshipId = slideIds[index].RelationshipId?.Value;
            if (relationshipId is null || presentationPart!.GetPartById(relationshipId) is not SlidePart slidePart) continue;
            var lines = slidePart.Slide.Descendants<A.Text>()
                .Select(text => text.Text.Trim()).Where(text => text.Length > 0).ToArray();
            if (lines.Length == 0) continue;
            var slideNumber = index + 1;
            chunks.Add(new ParsedChunk(
                string.Join('\n', new[] { $"【幻灯片 {slideNumber}】" }.Concat(lines)),
                sourceType,
                new ParsedMetadata(SlideNumber: slideNumber, RowCount: lines.Length)));
        }

        return chunks;
    }

    private static string[] ReadSpreadsheetRow(S.Row row, IReadOnlyList<string> sharedStrings)
    {
        var values = new SortedDictionary<int, string>();
        var nextColumn = 1;
        foreach (var cell in row.Elements<S.Cell>())
        {
            var column = ColumnNumber(cell.CellReference?.Value);
            if (column <= 0) column = nextColumn;
            values[column] = ReadCellValue(cell, sharedStrings);
            nextColumn = column + 1;
        }

        var width = values.Keys.DefaultIfEmpty(0).Max();
        return Enumerable.Range(1, width).Select(column => values.GetValueOrDefault(column, string.Empty)).ToArray();
    }

    private static string ReadCellValue(S.Cell cell, IReadOnlyList<string> sharedStrings)
    {
        if (cell.DataType?.Value == S.CellValues.InlineString) return cell.InlineString?.InnerText ?? string.Empty;
        var value = cell.CellValue?.Text ?? cell.InnerText;
        if (cell.DataType?.Value == S.CellValues.SharedString && int.TryParse(value, out var index))
        {
            return index >= 0 && index < sharedStrings.Count ? sharedStrings[index] : string.Empty;
        }
        if (cell.DataType?.Value == S.CellValues.Boolean) return value == "1" ? "TRUE" : "FALSE";
        return value;
    }

    private static int ColumnNumber(string? reference)
    {
        if (string.IsNullOrWhiteSpace(reference)) return 0;
        var result = 0;
        foreach (var character in reference.TakeWhile(char.IsLetter))
        {
            result = result * 26 + char.ToUpperInvariant(character) - 'A' + 1;
        }
        return result;
    }

    private static string CellReference(int column, int row)
    {
        var name = string.Empty;
        while (column > 0)
        {
            column--;
            name = (char)('A' + column % 26) + name;
            column /= 26;
        }
        return $"{name}{row}";
    }

    private sealed record ParsedChunk(string Content, string SourceType, ParsedMetadata Metadata);

    private sealed record ParsedMetadata(
        string? SheetName = null,
        string? TableRange = null,
        string[]? Headers = null,
        int? RowCount = null,
        int? ColCount = null,
        int? SlideNumber = null,
        string[][]? Rows = null);
}
