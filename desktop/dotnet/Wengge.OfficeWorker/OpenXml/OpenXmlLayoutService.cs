using DocumentFormat.OpenXml.Packaging;
using D = DocumentFormat.OpenXml.Drawing;
using S = DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.OpenXml;

internal sealed class OpenXmlLayoutService
{
    public object Inspect(string filePath, string? target)
    {
        var path = RequireFile(filePath);
        var extension = Path.GetExtension(path).ToLowerInvariant();
        var objects = extension switch
        {
            ".docx" or ".docm" => InspectWord(path),
            ".xlsx" or ".xlsm" => InspectSpreadsheet(path),
            ".pptx" or ".pptm" => InspectPresentation(path),
            _ => throw new OfficeWorkerException("unsupported_file", $"不支持的 Open XML 文件类型: {extension}"),
        };
        return new
        {
            engine = "openxml",
            operation = "inspectLayout",
            documentType = TypeName(extension),
            filePath = path,
            target,
            objectCount = objects.Count,
            objects,
        };
    }

    private static List<object> InspectWord(string path)
    {
        using var document = WordprocessingDocument.Open(path, false);
        return document.MainDocumentPart?.Document.Descendants<W.Text>()
            .Where(text => !string.IsNullOrEmpty(text.Text))
            .Select(text => (object)new { type = "text", partName = "word/document.xml", text = text.Text, textLength = text.Text.Length })
            .ToList() ?? [];
    }

    private static List<object> InspectSpreadsheet(string path)
    {
        using var document = SpreadsheetDocument.Open(path, false);
        var objects = new List<object>();
        var workbook = document.WorkbookPart;
        if (workbook?.SharedStringTablePart?.SharedStringTable is { } strings)
        {
            objects.AddRange(strings.Descendants<S.Text>().Where(text => !string.IsNullOrEmpty(text.Text)).Select(text =>
                (object)new { type = "text", partName = "xl/sharedStrings.xml", text = text.Text, textLength = text.Text.Length }));
        }
        foreach (var worksheet in workbook?.WorksheetParts ?? [])
        {
            var partName = worksheet.Uri.ToString().TrimStart('/');
            objects.AddRange(worksheet.Worksheet.Descendants<S.Text>().Where(text => !string.IsNullOrEmpty(text.Text)).Select(text =>
                (object)new { type = "text", partName, text = text.Text, textLength = text.Text.Length }));
        }
        return objects;
    }

    private static List<object> InspectPresentation(string path)
    {
        using var document = PresentationDocument.Open(path, false);
        var objects = new List<object>();
        foreach (var slide in document.PresentationPart?.SlideParts ?? [])
        {
            var partName = slide.Uri.ToString().TrimStart('/');
            objects.AddRange(slide.Slide.Descendants<D.Text>().Where(text => !string.IsNullOrEmpty(text.Text)).Select(text =>
                (object)new { type = "text", partName, text = text.Text, textLength = text.Text.Length }));
        }
        return objects;
    }

    private static string RequireFile(string path)
    {
        var fullPath = Path.GetFullPath(path);
        return File.Exists(fullPath) ? fullPath : throw new OfficeWorkerException("file_not_found", $"Office 文件不存在: {fullPath}");
    }

    private static string TypeName(string extension) => extension switch
    {
        ".docx" or ".docm" => "word",
        ".xlsx" or ".xlsm" => "spreadsheet",
        _ => "presentation",
    };
}
