using System.Text.Json;
using System.Text.RegularExpressions;
using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Office;

internal sealed class WordFormattingActionService(OfficeApplicationProvider applications)
{
    private static readonly TimeSpan HeadingPatternTimeout = TimeSpan.FromMilliseconds(100);
    private static readonly HashSet<string> Operations =
    [
        "applyHeadingStyles", "insertOrUpdateToc", "styleTables", "setHeaderFooter", "insertOrReplaceImage",
        "snapshot", "exportPdf", "inspectDocumentFormatting", "formatLongDocument",
    ];

    public static bool Supports(string operation) => Operations.Contains(operation);

    public object Execute(OfficeActionRequest request)
    {
        using var context = new WordActionContext(applications, request);
        if (request.Operation == "inspectDocumentFormatting")
            return OfficeActionResults.Done(request, "com", "已检查 Word 文档格式", Inspect(context));
        if (request.Operation is "snapshot" or "exportPdf")
        {
            var output = OutputPath(request, request.Operation == "snapshot" ? "preview.pdf" : "export.pdf");
            context.Document.ExportAsFixedFormat(output, 17);
            return OfficeActionResults.Done(request, "com", "已导出 Word PDF", new { outputPath = output }, [new OfficeChange("export", output, "已导出 Word PDF")], output);
        }
        var data = request.Operation switch
        {
            "applyHeadingStyles" => ApplyHeadingStyles(context, request),
            "insertOrUpdateToc" => UpdateToc(context, request),
            "styleTables" => StyleTables(context, request),
            "setHeaderFooter" => SetHeaderFooter(context, request),
            "insertOrReplaceImage" => InsertImage(context, request),
            "formatLongDocument" => FormatLongDocument(context, request),
            _ => throw new OfficeWorkerException("unsupported_operation", $"不支持的 Word 格式操作: {request.Operation}"),
        };
        context.Save(request);
        return OfficeActionResults.Done(request, "com", "已更新 Word 文档格式", data,
            [new OfficeChange("word-format", request.Target, "已更新 Word 文档格式")]);
    }

    private static object ApplyHeadingStyles(WordActionContext context, OfficeActionRequest request)
    {
        var prefix = request.StringParam("startsWith");
        var pattern = request.StringParam("pattern");
        var level = Math.Clamp(request.IntParam("level", 1), 1, 9);
        var applied = 0;
        object? paragraphs = null;
        try
        {
            paragraphs = context.Document.Paragraphs; dynamic paragraphsApi = paragraphs;
            for (var index = 1; index <= Convert.ToInt32(paragraphsApi.Count); index++)
            {
                object? paragraph = null;
                object? range = null;
                try
                {
                    paragraph = paragraphsApi.Item(index); range = ((dynamic)paragraph).Range; dynamic rangeApi = range;
                    var text = (Convert.ToString(rangeApi.Text) ?? string.Empty).Trim();
                    var matches = prefix.Length == 0 && pattern.Length == 0
                        || prefix.Length > 0 && text.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                        || MatchesHeadingPattern(text, pattern);
                    if (!matches) continue;
                    rangeApi.Style = -1 - level;
                    applied++;
                }
                finally { ComInterop.Release(range); ComInterop.Release(paragraph); }
            }
            return new { applied, level };
        }
        finally { ComInterop.Release(paragraphs); }
    }

    private static object UpdateToc(WordActionContext context, OfficeActionRequest request)
    {
        object? tables = null;
        object? range = null;
        object? toc = null;
        try
        {
            tables = context.Document.TablesOfContents; dynamic tablesApi = tables;
            if (Convert.ToInt32(tablesApi.Count) > 0)
            {
                for (var index = 1; index <= Convert.ToInt32(tablesApi.Count); index++)
                {
                    object? existing = null;
                    try { existing = tablesApi.Item(index); ((dynamic)existing).Update(); }
                    finally { ComInterop.Release(existing); }
                }
                return new { updated = Convert.ToInt32(tablesApi.Count), created = false };
            }
            var position = request.StringParam("position", "start");
            var offset = position == "end" ? Math.Max(0, Convert.ToInt32(context.Document.Content.End) - 1) : 0;
            range = context.Document.Range(offset, offset);
            toc = tablesApi.Add(range, true, request.IntParam("upperHeadingLevel", 1), request.IntParam("lowerHeadingLevel", 3));
            return new { updated = 0, created = true };
        }
        finally { ComInterop.Release(toc); ComInterop.Release(range); ComInterop.Release(tables); }
    }

    private static object StyleTables(WordActionContext context, OfficeActionRequest request)
    {
        var styled = 0;
        object? tables = null;
        try
        {
            tables = context.Document.Tables; dynamic tablesApi = tables;
            for (var index = 1; index <= Convert.ToInt32(tablesApi.Count); index++)
            {
                object? table = null;
                object? borders = null;
                object? rows = null;
                object? firstRow = null;
                object? range = null;
                object? shading = null;
                try
                {
                    table = tablesApi.Item(index); dynamic tableApi = table;
                    borders = tableApi.Borders; ((dynamic)borders).Enable = 1;
                    rows = tableApi.Rows;
                    if (Convert.ToInt32(((dynamic)rows).Count) > 0)
                    {
                        firstRow = ((dynamic)rows).Item(1); range = ((dynamic)firstRow).Range; dynamic rangeApi = range;
                        rangeApi.Bold = true; shading = rangeApi.Shading; ((dynamic)shading).BackgroundPatternColor = ExcelActionService.OleColor(request.StringParam("headerColor", "D9EAF7"));
                    }
                    try { tableApi.AutoFitBehavior(1); } catch { }
                    styled++;
                }
                finally { ComInterop.Release(shading); ComInterop.Release(range); ComInterop.Release(firstRow); ComInterop.Release(rows); ComInterop.Release(borders); ComInterop.Release(table); }
            }
            return new { styledTables = styled };
        }
        finally { ComInterop.Release(tables); }
    }

    private static object SetHeaderFooter(WordActionContext context, OfficeActionRequest request)
    {
        var text = request.StringParam("text");
        var footer = request.StringParam("kind") == "footer";
        var updated = 0;
        object? sections = null;
        try
        {
            sections = context.Document.Sections; dynamic sectionsApi = sections;
            for (var index = 1; index <= Convert.ToInt32(sectionsApi.Count); index++)
            {
                object? section = null; object? collection = null; object? item = null; object? range = null;
                try
                {
                    section = sectionsApi.Item(index); dynamic sectionApi = section;
                    collection = footer ? sectionApi.Footers : sectionApi.Headers;
                    item = ((dynamic)collection).Item(1); range = ((dynamic)item).Range; ((dynamic)range).Text = text; updated++;
                }
                finally { ComInterop.Release(range); ComInterop.Release(item); ComInterop.Release(collection); ComInterop.Release(section); }
            }
            return new { updatedSections = updated, kind = footer ? "footer" : "header" };
        }
        finally { ComInterop.Release(sections); }
    }

    private static object InsertImage(WordActionContext context, OfficeActionRequest request)
    {
        var imagePath = Path.GetFullPath(request.StringParam("imagePath"));
        if (!File.Exists(imagePath)) throw new OfficeWorkerException("file_not_found", $"图片不存在: {imagePath}");
        object? range = null; object? images = null; object? image = null;
        try
        {
            var bookmark = request.StringParam("bookmark");
            if (bookmark.Length > 0)
            {
                object? bookmarks = null; object? item = null;
                try { bookmarks = context.Document.Bookmarks; item = ((dynamic)bookmarks).Item(bookmark); range = ((dynamic)item).Range; }
                finally { ComInterop.Release(item); ComInterop.Release(bookmarks); }
            }
            range ??= context.Document.Range(Math.Max(0, Convert.ToInt32(context.Document.Content.End) - 1), Math.Max(0, Convert.ToInt32(context.Document.Content.End) - 1));
            images = context.Document.InlineShapes;
            image = ((dynamic)images).AddPicture(imagePath, false, true, range);
            if (request.DoubleParam("width") > 0) ((dynamic)image).Width = request.DoubleParam("width");
            if (request.DoubleParam("height") > 0) ((dynamic)image).Height = request.DoubleParam("height");
            return new { inserted = true, imagePath };
        }
        finally { ComInterop.Release(image); ComInterop.Release(images); ComInterop.Release(range); }
    }

    private static object FormatLongDocument(WordActionContext context, OfficeActionRequest request)
    {
        var normalFont = request.StringParam("fontName", "微软雅黑");
        var fontSize = request.DoubleParam("fontSize", 10.5);
        object? content = null; object? contentFont = null;
        try
        {
            content = context.Document.Content;
            contentFont = ((dynamic)content).Font;
            ((dynamic)contentFont).Name = normalFont;
            ((dynamic)contentFont).Size = fontSize;
        }
        finally { ComInterop.Release(contentFont); ComInterop.Release(content); }
        var headingResult = request.BoolParam("autoDetectHeadings")
            ? AutoDetectHeadings(context)
            : ApplyHeadingStyles(context, request);
        var tableResult = StyleTables(context, request);
        var layoutResult = ApplyLongDocumentLayout(context, request);
        object? tocResult = null;
        if (request.StringParam("toc") is "create" or "update") tocResult = UpdateToc(context, request);
        object? fields = null;
        try { fields = context.Document.Fields; ((dynamic)fields).Update(); }
        catch { }
        finally { ComInterop.Release(fields); }
        return new { fontName = normalFont, fontSize, headingResult, tableResult, layoutResult, tocResult };
    }

    private static object AutoDetectHeadings(WordActionContext context)
    {
        var applied = new List<object>();
        object? paragraphs = null;
        try
        {
            paragraphs = context.Document.Paragraphs; dynamic paragraphsApi = paragraphs;
            for (var index = 1; index <= Convert.ToInt32(paragraphsApi.Count); index++)
            {
                object? paragraph = null; object? range = null;
                try
                {
                    paragraph = paragraphsApi.Item(index); range = ((dynamic)paragraph).Range; dynamic rangeApi = range;
                    var text = (Convert.ToString(rangeApi.Text) ?? string.Empty).Trim();
                    var level = Regex.IsMatch(text, @"^(第[一二三四五六七八九十百]+[章节]|[一二三四五六七八九十]+、|\d+\s*[、.])") ? 1
                        : Regex.IsMatch(text, @"^(（[一二三四五六七八九十]+）|\d+\.\d+)") ? 2
                        : 0;
                    if (level == 0) continue;
                    rangeApi.Style = -1 - level;
                    applied.Add(new { index, level, text });
                }
                finally { ComInterop.Release(range); ComInterop.Release(paragraph); }
            }
            return new { applied = applied.Count, headings = applied };
        }
        finally { ComInterop.Release(paragraphs); }
    }

    private static object ApplyLongDocumentLayout(WordActionContext context, OfficeActionRequest request)
    {
        var margins = request.Param("margins");
        var headerFooter = request.Param("headerFooter");
        var pageNumbers = request.BoolParam("pageNumbers");
        var updatedSections = 0;
        object? sections = null;
        try
        {
            sections = context.Document.Sections; dynamic sectionsApi = sections;
            for (var index = 1; index <= Convert.ToInt32(sectionsApi.Count); index++)
            {
                object? section = null; object? setup = null; object? headers = null; object? header = null;
                object? headerRange = null; object? footers = null; object? footer = null; object? footerRange = null; object? numbers = null;
                try
                {
                    section = sectionsApi.Item(index); dynamic sectionApi = section;
                    if (margins.ValueKind == JsonValueKind.Object)
                    {
                        setup = sectionApi.PageSetup; dynamic setupApi = setup;
                        SetMargin(setupApi, margins, "top", "TopMargin");
                        SetMargin(setupApi, margins, "bottom", "BottomMargin");
                        SetMargin(setupApi, margins, "left", "LeftMargin");
                        SetMargin(setupApi, margins, "right", "RightMargin");
                    }
                    if (headerFooter.ValueKind == JsonValueKind.Object && headerFooter.TryGetProperty("header", out var headerText) && headerText.ValueKind == JsonValueKind.String)
                    {
                        headers = sectionApi.Headers; header = ((dynamic)headers).Item(1); headerRange = ((dynamic)header).Range; ((dynamic)headerRange).Text = headerText.GetString();
                    }
                    if (headerFooter.ValueKind == JsonValueKind.Object && headerFooter.TryGetProperty("footer", out var footerText) && footerText.ValueKind == JsonValueKind.String)
                    {
                        footers = sectionApi.Footers; footer = ((dynamic)footers).Item(1); footerRange = ((dynamic)footer).Range; ((dynamic)footerRange).Text = footerText.GetString();
                    }
                    if (pageNumbers)
                    {
                        footers ??= sectionApi.Footers; footer ??= ((dynamic)footers).Item(1); numbers = ((dynamic)footer).PageNumbers; ((dynamic)numbers).Add(1, true);
                    }
                    updatedSections++;
                }
                finally
                {
                    ComInterop.Release(numbers); ComInterop.Release(footerRange); ComInterop.Release(footer); ComInterop.Release(footers);
                    ComInterop.Release(headerRange); ComInterop.Release(header); ComInterop.Release(headers); ComInterop.Release(setup); ComInterop.Release(section);
                }
            }
            return new { updatedSections, pageNumbers };
        }
        finally { ComInterop.Release(sections); }
    }

    private static void SetMargin(dynamic setup, JsonElement margins, string propertyName, string comProperty)
    {
        if (!margins.TryGetProperty(propertyName, out var value) || !value.TryGetDouble(out var centimeters)) return;
        var points = centimeters * 72d / 2.54d;
        if (comProperty == "TopMargin") setup.TopMargin = points;
        else if (comProperty == "BottomMargin") setup.BottomMargin = points;
        else if (comProperty == "LeftMargin") setup.LeftMargin = points;
        else setup.RightMargin = points;
    }

    private static object Inspect(WordActionContext context)
    {
        var styles = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        object? paragraphs = null;
        try
        {
            paragraphs = context.Document.Paragraphs; dynamic paragraphsApi = paragraphs;
            for (var index = 1; index <= Convert.ToInt32(paragraphsApi.Count); index++)
            {
                object? paragraph = null; object? range = null; object? style = null;
                try
                {
                    paragraph = paragraphsApi.Item(index); range = ((dynamic)paragraph).Range; style = ((dynamic)range).Style;
                    string name = Convert.ToString(((dynamic)style).NameLocal) ?? "Unknown";
                    styles[name] = (styles.TryGetValue(name, out var count) ? count : 0) + 1;
                }
                catch { }
                finally { ComInterop.Release(style); ComInterop.Release(range); ComInterop.Release(paragraph); }
            }
            return new { progId = context.ProgId, paragraphCount = Convert.ToInt32(paragraphsApi.Count), styles = styles.Select(pair => new { name = pair.Key, count = pair.Value }) };
        }
        finally { ComInterop.Release(paragraphs); }
    }

    internal static bool MatchesHeadingPattern(string text, string pattern)
    {
        if (pattern.Length == 0) return false;
        try
        {
            return Regex.IsMatch(text, pattern, RegexOptions.CultureInvariant, HeadingPatternTimeout);
        }
        catch (RegexMatchTimeoutException exception)
        {
            throw new OfficeWorkerException("invalid_params", "Word 标题匹配正则执行超时", inner: exception);
        }
        catch (ArgumentException exception)
        {
            throw new OfficeWorkerException("invalid_params", "Word 标题匹配正则无效", inner: exception);
        }
    }

    private static string OutputPath(OfficeActionRequest request, string fallback)
    {
        var output = request.OutputPath ?? Path.Combine(Path.GetDirectoryName(request.FilePath) ?? Environment.CurrentDirectory, fallback);
        output = Path.GetFullPath(output); Directory.CreateDirectory(Path.GetDirectoryName(output) ?? Environment.CurrentDirectory); return output;
    }
}
