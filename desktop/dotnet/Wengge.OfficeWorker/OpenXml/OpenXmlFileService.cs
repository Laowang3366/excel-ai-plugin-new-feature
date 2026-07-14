using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Wordprocessing;
using D = DocumentFormat.OpenXml.Drawing;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.OpenXml;

internal sealed class OpenXmlFileService
{
    public object Inspect(string filePath)
    {
        var fullPath = RequireFile(filePath);
        var extension = Path.GetExtension(fullPath).ToLowerInvariant();
        var parts = extension switch
        {
            ".docx" or ".docm" => InspectWord(fullPath),
            ".xlsx" or ".xlsm" => InspectSpreadsheet(fullPath),
            ".pptx" or ".pptm" => InspectPresentation(fullPath),
            _ => throw Unsupported(extension),
        };
        var text = string.Join('\n', parts.Select(part => part.Text));
        return new
        {
            engine = "openxml",
            operation = "inspect",
            documentType = extension[1..],
            filePath = fullPath,
            textPartCount = parts.Count,
            textCharCount = text.Length,
            textPreview = text[..Math.Min(text.Length, 4_000)],
            textParts = parts.Select(part => new { partName = part.Name, text = part.Text, textLength = part.Text.Length }),
        };
    }

    public object ReplaceText(string filePath, string findText, string replaceText, string? outputPath, bool matchCase)
    {
        var sourcePath = RequireFile(filePath);
        if (string.IsNullOrEmpty(findText))
        {
            throw new OfficeWorkerException("invalid_params", "findText 不能为空");
        }

        var targetPath = string.IsNullOrWhiteSpace(outputPath) ? DefaultOutputPath(sourcePath) : Path.GetFullPath(outputPath);
        Directory.CreateDirectory(Path.GetDirectoryName(targetPath) ?? Environment.CurrentDirectory);
        if (!SamePath(sourcePath, targetPath)) File.Copy(sourcePath, targetPath, overwrite: true);
        var extension = Path.GetExtension(targetPath).ToLowerInvariant();
        var changedParts = extension switch
        {
            ".docx" or ".docm" => ReplaceWord(targetPath, findText, replaceText, matchCase),
            ".xlsx" or ".xlsm" => ReplaceSpreadsheet(targetPath, findText, replaceText, matchCase),
            ".pptx" or ".pptm" => ReplacePresentation(targetPath, findText, replaceText, matchCase),
            _ => throw Unsupported(extension),
        };
        return new
        {
            engine = "openxml",
            operation = "replaceText",
            filePath = sourcePath,
            outputPath = targetPath,
            findText,
            replaceText,
            replacements = changedParts.Sum(part => part.Count),
            changedParts = changedParts.Select(part => new { partName = part.Name, replacements = part.Count }),
        };
    }

    private static List<TextPart> InspectWord(string path)
    {
        using var document = WordprocessingDocument.Open(path, false);
        var main = document.MainDocumentPart;
        if (main is null) return [];

        var parts = new List<TextPart>
        {
            new("word/document.xml", WordText(main.Document)),
        };
        parts.AddRange(main.HeaderParts
            .Select(part => new TextPart(PartName(part), WordText(part.Header)))
            .Where(part => part.Text.Length > 0));
        parts.AddRange(main.FooterParts
            .Select(part => new TextPart(PartName(part), WordText(part.Footer)))
            .Where(part => part.Text.Length > 0));
        return parts;
    }

    private static List<TextPart> InspectSpreadsheet(string path)
    {
        using var document = SpreadsheetDocument.Open(path, false);
        var workbook = document.WorkbookPart;
        if (workbook is null)
        {
            return [];
        }

        var parts = new List<TextPart>();
        var sharedStrings = workbook.SharedStringTablePart?.SharedStringTable;
        if (sharedStrings is not null)
        {
            parts.Add(new TextPart("xl/sharedStrings.xml", string.Join('\n', sharedStrings.Elements<S.SharedStringItem>().Select(item => item.InnerText))));
        }

        foreach (var worksheet in workbook.WorksheetParts)
        {
            parts.Add(new TextPart(worksheet.Uri.ToString().TrimStart('/'), string.Join('\n', worksheet.Worksheet.Descendants<S.CellValue>().Select(value => value.Text))));
        }

        return parts.Where(part => part.Text.Length > 0).ToList();
    }

    private static List<TextPart> InspectPresentation(string path)
    {
        using var document = PresentationDocument.Open(path, false);
        var presentation = document.PresentationPart;
        return presentation is null
            ? []
            : presentation.SlideParts.Select(slide =>
                new TextPart(slide.Uri.ToString().TrimStart('/'), string.Join(string.Empty, slide.Slide.Descendants<D.Text>().Select(text => text.Text))))
                .Where(part => part.Text.Length > 0)
                .ToList();
    }

    private static List<ChangedPart> ReplaceWord(string path, string find, string replacement, bool matchCase)
    {
        using var document = WordprocessingDocument.Open(path, true);
        var main = document.MainDocumentPart;
        if (main is null)
        {
            return [];
        }

        var changed = new List<ChangedPart>();
        ReplaceWordPart(main.Document, "word/document.xml", main.Document.Save, changed, find, replacement, matchCase);
        foreach (var header in main.HeaderParts)
            ReplaceWordPart(header.Header, PartName(header), header.Header.Save, changed, find, replacement, matchCase);
        foreach (var footer in main.FooterParts)
            ReplaceWordPart(footer.Footer, PartName(footer), footer.Footer.Save, changed, find, replacement, matchCase);
        return changed;
    }

    private static void ReplaceWordPart(
        OpenXmlElement root,
        string partName,
        Action save,
        ICollection<ChangedPart> changed,
        string find,
        string replacement,
        bool matchCase)
    {
        var count = ReplaceNodes(root.Descendants<W.Text>(), node => node.Text, (node, value) => node.Text = value, find, replacement, matchCase);
        if (count == 0) return;
        save();
        changed.Add(new ChangedPart(partName, count));
    }

    private static List<ChangedPart> ReplaceSpreadsheet(string path, string find, string replacement, bool matchCase)
    {
        using var document = SpreadsheetDocument.Open(path, true);
        var workbook = document.WorkbookPart;
        if (workbook is null)
        {
            return [];
        }

        var changed = new List<ChangedPart>();
        var sharedStrings = workbook.SharedStringTablePart?.SharedStringTable;
        if (sharedStrings is not null)
        {
            var count = ReplaceNodes(sharedStrings.Descendants<S.Text>(), node => node.Text, (node, value) => node.Text = value, find, replacement, matchCase);
            if (count > 0)
            {
                sharedStrings.Save();
                changed.Add(new ChangedPart("xl/sharedStrings.xml", count));
            }
        }

        foreach (var worksheet in workbook.WorksheetParts)
        {
            var count = ReplaceNodes(worksheet.Worksheet.Descendants<S.Text>(), node => node.Text, (node, value) => node.Text = value, find, replacement, matchCase);
            if (count > 0)
            {
                worksheet.Worksheet.Save();
                changed.Add(new ChangedPart(worksheet.Uri.ToString().TrimStart('/'), count));
            }
        }

        return changed;
    }

    private static List<ChangedPart> ReplacePresentation(string path, string find, string replacement, bool matchCase)
    {
        using var document = PresentationDocument.Open(path, true);
        var presentation = document.PresentationPart;
        if (presentation is null)
        {
            return [];
        }

        var changed = new List<ChangedPart>();
        foreach (var slide in presentation.SlideParts)
        {
            var count = ReplaceNodes(slide.Slide.Descendants<D.Text>(), node => node.Text, (node, value) => node.Text = value, find, replacement, matchCase);
            if (count > 0)
            {
                slide.Slide.Save();
                changed.Add(new ChangedPart(slide.Uri.ToString().TrimStart('/'), count));
            }
        }

        return changed;
    }

    private static int ReplaceNodes<T>(
        IEnumerable<T> nodes,
        Func<T, string> read,
        Action<T, string> write,
        string find,
        string replacement,
        bool matchCase)
    {
        var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var total = 0;
        foreach (var node in nodes)
        {
            var source = read(node);
            var next = ReplaceAll(source, find, replacement, comparison, out var count);
            if (count == 0) continue;
            write(node, next);
            total += count;
        }

        return total;
    }

    private static string ReplaceAll(string source, string find, string replacement, StringComparison comparison, out int count)
    {
        count = 0;
        var start = 0;
        var writer = new System.Text.StringBuilder(source.Length);
        while (true)
        {
            var index = source.IndexOf(find, start, comparison);
            if (index < 0) break;
            writer.Append(source, start, index - start).Append(replacement);
            start = index + find.Length;
            count++;
        }

        return count == 0 ? source : writer.Append(source, start, source.Length - start).ToString();
    }

    private static string RequireFile(string path)
    {
        var fullPath = Path.GetFullPath(path);
        return File.Exists(fullPath)
            ? fullPath
            : throw new OfficeWorkerException("file_not_found", $"Office 文件不存在: {fullPath}");
    }

    private static string DefaultOutputPath(string path) =>
        Path.Combine(Path.GetDirectoryName(path) ?? Environment.CurrentDirectory, $"{Path.GetFileNameWithoutExtension(path)}-edited{Path.GetExtension(path)}");

    private static bool SamePath(string left, string right) =>
        string.Equals(Path.GetFullPath(left), Path.GetFullPath(right), StringComparison.OrdinalIgnoreCase);

    private static string WordText(OpenXmlElement root) =>
        string.Join(string.Empty, root.Descendants<W.Text>().Select(text => text.Text));

    private static string PartName(OpenXmlPart part) => part.Uri.ToString().TrimStart('/');

    private static OfficeWorkerException Unsupported(string extension) =>
        new("unsupported_file", $"不支持的 Open XML 文件类型: {extension}");

    private sealed record TextPart(string Name, string Text);

    private sealed record ChangedPart(string Name, int Count);
}
