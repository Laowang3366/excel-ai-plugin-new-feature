using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wengge.OfficeWorker.Office;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.OpenXml;

internal sealed class OpenXmlWordActionService(OpenXmlTableService tables)
{
    private static readonly HashSet<string> Operations =
    ["createDocument", "applyHeadingStyles", "styleTables", "setHeaderFooter", "insertOrUpdateToc", "insertOrReplaceImage"];

    public static bool Supports(string operation) => Operations.Contains(operation);

    public object Execute(OfficeActionRequest request) => request.Operation switch
    {
        "createDocument" => CreateDocument(request),
        "applyHeadingStyles" => ApplyHeadingStyles(request),
        "styleTables" => StyleTables(request),
        "setHeaderFooter" => SetHeaderFooter(request),
        "insertOrUpdateToc" or "insertOrReplaceImage" => OfficeActionResults.NeedsCom(request, $"{request.Operation} 需要 Word 刷新字段或维护媒体关系，需要 COM 执行"),
        _ => throw new OfficeWorkerException("unsupported_operation", $"不支持的 Word Open XML 操作: {request.Operation}"),
    };

    private static object CreateDocument(OfficeActionRequest request)
    {
        var output = RequiredOutput(request);
        Directory.CreateDirectory(Path.GetDirectoryName(output) ?? Environment.CurrentDirectory);
        if (File.Exists(output)) File.Delete(output);
        var title = request.StringParam("title");
        var paragraphs = Paragraphs(request.Param("paragraphs"));
        if (paragraphs.Count == 0) paragraphs = Paragraphs(request.Param("text"));
        if (paragraphs.Count == 0) paragraphs = Paragraphs(request.Param("body"));
        using var document = WordprocessingDocument.Create(output, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        main.Document = new W.Document();
        var body = main.Document.AppendChild(new W.Body());
        if (title.Length > 0) body.Append(StyledParagraph(title, "Title"));
        foreach (var paragraph in paragraphs) body.Append(StyledParagraph(paragraph, null));
        body.Append(new W.SectionProperties(
            new W.PageSize { Width = 12240U, Height = 15840U },
            new W.PageMargin { Top = 1440, Right = 1440U, Bottom = 1440, Left = 1440U, Header = 720U, Footer = 720U, Gutter = 0U }));
        var styles = main.AddNewPart<StyleDefinitionsPart>();
        styles.Styles = BaseStyles();
        styles.Styles.Save();
        main.Document.Save();
        return Done(request, "已使用 .NET Open XML 创建 Word 文档", output, ["word/document.xml", "word/styles.xml"],
            new { title, paragraphCount = paragraphs.Count });
    }

    private static object ApplyHeadingStyles(OfficeActionRequest request)
    {
        var output = PrepareCopy(request, "advanced");
        using var document = WordprocessingDocument.Open(output, true);
        var main = document.MainDocumentPart ?? throw new OfficeWorkerException("invalid_file", "Word 文档缺少正文部件");
        var startsWith = request.StringParam("startsWith");
        var level = Math.Clamp(request.IntParam("level", 1), 1, 9);
        var changed = 0;
        foreach (var paragraph in main.Document.Descendants<W.Paragraph>())
        {
            var text = paragraph.InnerText;
            if (text.Length == 0 || startsWith.Length > 0 && !text.StartsWith(startsWith, StringComparison.Ordinal)) continue;
            paragraph.ParagraphProperties ??= new W.ParagraphProperties();
            paragraph.ParagraphProperties.ParagraphStyleId = new W.ParagraphStyleId { Val = $"Heading{level}" };
            changed++;
        }
        EnsureHeadingStyles(main, level);
        main.Document.Save();
        return Done(request, $"已为 {changed} 个 Word 段落应用标题 {level}", output, ["word/document.xml", "word/styles.xml"], new { changed, level });
    }

    private object StyleTables(OfficeActionRequest request)
    {
        var source = RequireFile(request.FilePath);
        var data = tables.ApplyStyle(source, request.StringParam("style", "professional"), request.OutputPath, request.Target);
        var output = ReadProperty(data, "outputPath") ?? request.OutputPath ?? source;
        return OfficeActionResults.Done(request, "openxml", "已应用 Word 表格样式", data,
            [new OfficeChange("openxml-part", "word/document.xml", "已更新 Word 表格样式")], output);
    }

    private static object SetHeaderFooter(OfficeActionRequest request)
    {
        var output = PrepareCopy(request, "advanced");
        var footer = string.Equals(request.StringParam("kind"), "footer", StringComparison.OrdinalIgnoreCase);
        var text = request.StringParam("text");
        using var document = WordprocessingDocument.Open(output, true);
        var main = document.MainDocumentPart ?? throw new OfficeWorkerException("invalid_file", "Word 文档缺少正文部件");
        var section = main.Document.Body?.Elements<W.SectionProperties>().LastOrDefault();
        if (section is null)
        {
            section = new W.SectionProperties();
            main.Document.Body ??= new W.Body();
            main.Document.Body.Append(section);
        }
        string relationshipId;
        string partName;
        if (footer)
        {
            var part = main.AddNewPart<FooterPart>();
            part.Footer = new W.Footer(new W.Paragraph(new W.Run(new W.Text(text))));
            part.Footer.Save();
            relationshipId = main.GetIdOfPart(part);
            partName = part.Uri.ToString().TrimStart('/');
            var reference = new W.FooterReference { Type = W.HeaderFooterValues.Default, Id = relationshipId };
            var lastHeader = section.Elements<W.HeaderReference>().LastOrDefault();
            if (lastHeader is null) section.PrependChild(reference); else section.InsertAfter(reference, lastHeader);
        }
        else
        {
            var part = main.AddNewPart<HeaderPart>();
            part.Header = new W.Header(new W.Paragraph(new W.Run(new W.Text(text))));
            part.Header.Save();
            relationshipId = main.GetIdOfPart(part);
            partName = part.Uri.ToString().TrimStart('/');
            var reference = new W.HeaderReference { Type = W.HeaderFooterValues.Default, Id = relationshipId };
            var firstFooter = section.Elements<W.FooterReference>().FirstOrDefault();
            if (firstFooter is null) section.PrependChild(reference); else section.InsertBefore(reference, firstFooter);
        }
        main.Document.Save();
        return Done(request, footer ? "已设置 Word 页脚" : "已设置 Word 页眉", output,
            [partName, "word/document.xml", "word/_rels/document.xml.rels", "[Content_Types].xml"],
            new { kind = footer ? "footer" : "header", text, relationshipId });
    }

    private static W.Paragraph StyledParagraph(string text, string? style)
    {
        var paragraph = new W.Paragraph();
        if (!string.IsNullOrWhiteSpace(style)) paragraph.ParagraphProperties = new W.ParagraphProperties(new W.ParagraphStyleId { Val = style });
        paragraph.Append(new W.Run(new W.Text(text) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }));
        return paragraph;
    }

    private static W.Styles BaseStyles()
    {
        var styles = new W.Styles();
        styles.Append(new W.Style(
            new W.StyleName { Val = "Normal" },
            new W.BasedOn { Val = "Normal" },
            new W.UIPriority { Val = 0 },
            new W.PrimaryStyle()) { Type = W.StyleValues.Paragraph, StyleId = "Normal", Default = true });
        styles.Append(new W.Style(
            new W.StyleName { Val = "Title" },
            new W.BasedOn { Val = "Normal" },
            new W.NextParagraphStyle { Val = "Normal" },
            new W.StyleRunProperties(new W.Bold(), new W.FontSize { Val = "32" })) { Type = W.StyleValues.Paragraph, StyleId = "Title" });
        for (var level = 1; level <= 9; level++) styles.Append(HeadingStyle(level));
        return styles;
    }

    private static W.Style HeadingStyle(int level) => new(
        new W.StyleName { Val = $"heading {level}" },
        new W.BasedOn { Val = "Normal" },
        new W.NextParagraphStyle { Val = "Normal" },
        new W.UIPriority { Val = level + 8 },
        new W.PrimaryStyle(),
        new W.StyleParagraphProperties(new W.OutlineLevel { Val = level - 1 }),
        new W.StyleRunProperties(new W.Bold(), new W.FontSize { Val = level == 1 ? "28" : level == 2 ? "26" : "24" }))
    { Type = W.StyleValues.Paragraph, StyleId = $"Heading{level}" };

    private static void EnsureHeadingStyles(MainDocumentPart main, int level)
    {
        var part = main.StyleDefinitionsPart ?? main.AddNewPart<StyleDefinitionsPart>();
        part.Styles ??= BaseStyles();
        if (!part.Styles.Elements<W.Style>().Any(style => style.StyleId?.Value == $"Heading{level}")) part.Styles.Append(HeadingStyle(level));
        part.Styles.Save();
    }

    private static List<string> Paragraphs(JsonElement value)
    {
        if (value.ValueKind == JsonValueKind.Array)
            return value.EnumerateArray().Select(item => item.ValueKind == JsonValueKind.String ? item.GetString() ?? string.Empty : item.ToString()).Where(text => text.Length > 0).ToList();
        if (value.ValueKind == JsonValueKind.String)
            return (value.GetString() ?? string.Empty).Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n').Where(text => text.Length > 0).ToList();
        return [];
    }

    private static string PrepareCopy(OfficeActionRequest request, string suffix)
    {
        var source = RequireFile(request.FilePath);
        var output = !string.IsNullOrWhiteSpace(request.OutputPath) ? Path.GetFullPath(request.OutputPath)
            : Path.Combine(Path.GetDirectoryName(source) ?? Environment.CurrentDirectory, $"{Path.GetFileNameWithoutExtension(source)}-{suffix}{Path.GetExtension(source)}");
        Directory.CreateDirectory(Path.GetDirectoryName(output) ?? Environment.CurrentDirectory);
        if (!string.Equals(source, output, StringComparison.OrdinalIgnoreCase)) File.Copy(source, output, true);
        return output;
    }

    private static string RequiredOutput(OfficeActionRequest request)
    {
        var path = request.OutputPath ?? request.FilePath;
        return string.IsNullOrWhiteSpace(path) ? throw new OfficeWorkerException("invalid_params", "createDocument 需要 filePath 或 outputPath") : Path.GetFullPath(path);
    }

    private static string RequireFile(string? path)
    {
        if (string.IsNullOrWhiteSpace(path)) throw new OfficeWorkerException("invalid_params", "缺少 filePath");
        var fullPath = Path.GetFullPath(path);
        return File.Exists(fullPath) ? fullPath : throw new OfficeWorkerException("file_not_found", $"Office 文件不存在: {fullPath}");
    }

    private static object Done(OfficeActionRequest request, string summary, string output, IEnumerable<string> parts, object? data = null) =>
        OfficeActionResults.Done(request, "openxml", summary, data ?? new { outputPath = output, changedParts = parts.ToArray() },
            parts.Select(part => new OfficeChange("openxml-part", part, $"已更新 {part}")), output);

    private static string? ReadProperty(object value, string property) => value.GetType().GetProperty(property)?.GetValue(value) as string;
}
