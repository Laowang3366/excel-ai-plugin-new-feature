using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Word;

internal sealed class WordService(OfficeApplicationProvider applications)
{
    private static readonly string[] ProgIds = ["Word.Application", "Kwps.Application", "Wps.Application"];
    private string? preferredProgId;
    private string? activeDocumentPath;

    public object DetectStatus()
    {
        using var handle = applications.TryGetActive(OrderedProgIds());
        if (handle is null) return new { connected = false, host = "unknown" };
        preferredProgId = handle.ProgId;
        dynamic app = handle.Application;
        string? documentName = null;
        try { documentName = Convert.ToString(app.ActiveDocument?.Name); } catch { }
        return new
        {
            connected = true,
            host = handle.ProgId.Contains("wps", StringComparison.OrdinalIgnoreCase) ? "wps" : "word",
            version = Convert.ToString(app.Version),
            documentName,
        };
    }

    public object Open(string filePath)
    {
        var fullPath = Path.GetFullPath(filePath);
        if (!File.Exists(fullPath)) throw new OfficeWorkerException("file_not_found", $"Word 文档不存在: {fullPath}");
        using var handle = applications.GetOrCreate(OrderedProgIds(), "未找到可用的 Word/WPS 文字 COM 应用");
        dynamic app = handle.Application;
        object? documents = null;
        object? document = null;
        try
        {
            app.Visible = true;
            documents = app.Documents;
            dynamic documentsApi = documents;
            document = documentsApi.Open(fullPath);
            dynamic documentApi = document;
            documentApi.Activate();
            preferredProgId = handle.ProgId;
            activeDocumentPath = Convert.ToString(documentApi.FullName) ?? fullPath;
            return new { success = true, documentName = Convert.ToString(documentApi.Name) };
        }
        finally
        {
            ComInterop.Release(document);
            ComInterop.Release(documents);
        }
    }

    public object Inspect()
    {
        return WithDocument((app, document) =>
        {
            dynamic doc = document;
            object? paragraphs = null;
            object? tables = null;
            object? sections = null;
            object? words = null;
            object? characters = null;
            try
            {
                paragraphs = doc.Paragraphs;
                tables = doc.Tables;
                sections = doc.Sections;
                words = doc.Words;
                characters = doc.Characters;
                dynamic paragraphsApi = paragraphs;
                var preview = new List<string>();
                var count = Math.Min(8, Convert.ToInt32(paragraphsApi.Count));
                for (var index = 1; index <= count; index++)
                {
                    object? paragraph = null;
                    object? range = null;
                    try
                    {
                        paragraph = paragraphsApi.Item(index);
                        dynamic paragraphApi = paragraph;
                        range = paragraphApi.Range;
                        dynamic rangeApi = range;
                        preview.Add((Convert.ToString(rangeApi.Text) ?? string.Empty).Trim());
                    }
                    finally
                    {
                        ComInterop.Release(range);
                        ComInterop.Release(paragraph);
                    }
                }

                return new
                {
                    app = Convert.ToString(app.Name),
                    progId = preferredProgId,
                    name = Convert.ToString(doc.Name),
                    path = Convert.ToString(doc.FullName),
                    paragraphs = Convert.ToInt32(((dynamic)paragraphs).Count),
                    tables = Convert.ToInt32(((dynamic)tables).Count),
                    sections = Convert.ToInt32(((dynamic)sections).Count),
                    words = Convert.ToInt32(((dynamic)words).Count),
                    characters = Convert.ToInt32(((dynamic)characters).Count),
                    preview,
                };
            }
            finally
            {
                ComInterop.Release(characters);
                ComInterop.Release(words);
                ComInterop.Release(sections);
                ComInterop.Release(tables);
                ComInterop.Release(paragraphs);
            }
        });
    }

    public object ReadText(int maxChars)
    {
        return WithDocument((_, document) =>
        {
            dynamic doc = document;
            object? content = null;
            try
            {
                content = doc.Content;
                dynamic contentApi = content;
                var text = Convert.ToString(contentApi.Text) ?? string.Empty;
                var limit = Math.Max(1, maxChars);
                return new
                {
                    name = Convert.ToString(doc.Name),
                    path = Convert.ToString(doc.FullName),
                    text = text[..Math.Min(text.Length, limit)],
                    charCount = text.Length,
                    truncated = text.Length > limit,
                };
            }
            finally
            {
                ComInterop.Release(content);
            }
        });
    }

    public object InsertText(string text, string position)
    {
        return WithDocument((app, document) =>
        {
            dynamic doc = document;
            object? range = null;
            object? selection = null;
            try
            {
                if (position == "selection")
                {
                    selection = app.Selection;
                    dynamic selectionApi = selection;
                    selectionApi.TypeText(text);
                }
                else
                {
                    var offset = position == "start" ? 0 : Math.Max(0, Convert.ToInt32(doc.Content.End) - 1);
                    range = doc.Range(offset, offset);
                    dynamic rangeApi = range;
                    if (position == "start") rangeApi.InsertBefore(text); else rangeApi.InsertAfter(text);
                }

                return new { inserted = true, position, characters = text.Length };
            }
            finally
            {
                ComInterop.Release(selection);
                ComInterop.Release(range);
            }
        });
    }

    public object InsertHeading(string text, int level, string position)
    {
        return WithDocument((app, document) =>
        {
            dynamic doc = document;
            object? range = null;
            object? headingRange = null;
            object? style = null;
            try
            {
                var start = position == "start" ? 0 : Math.Max(0, Convert.ToInt32(doc.Content.End) - 1);
                if (position == "selection")
                {
                    object? selectionRange = app.Selection.Range;
                    try
                    {
                        dynamic selectionRangeApi = selectionRange!;
                        start = Convert.ToInt32(selectionRangeApi.Start);
                        app.Selection.TypeText(text + Environment.NewLine);
                    }
                    finally
                    {
                        ComInterop.Release(selectionRange);
                    }
                }
                else
                {
                    range = doc.Range(start, start);
                    dynamic rangeApi = range;
                    if (position == "start") rangeApi.InsertBefore(text + Environment.NewLine);
                    else rangeApi.InsertAfter(text + Environment.NewLine);
                }

                headingRange = doc.Range(start, start + text.Length);
                dynamic headingRangeApi = headingRange;
                style = doc.Styles.Item(-1 - Math.Clamp(level, 1, 9));
                headingRangeApi.Style = style;
                return new { inserted = true, position, level = Math.Clamp(level, 1, 9), characters = text.Length };
            }
            finally
            {
                ComInterop.Release(style);
                ComInterop.Release(headingRange);
                ComInterop.Release(range);
            }
        });
    }

    public object ReplaceText(string findText, string replaceText, bool matchCase)
    {
        return WithDocument((_, document) =>
        {
            dynamic doc = document;
            var replacements = 0;
            object? content = null;
            try
            {
                content = doc.Content;
                while (true)
                {
                    dynamic range = content;
                    object? find = null;
                    try
                    {
                        find = range.Find;
                        dynamic findApi = find;
                        findApi.ClearFormatting();
                        findApi.Text = findText;
                        findApi.Forward = true;
                        findApi.Wrap = 0;
                        findApi.MatchCase = matchCase;
                        if (!Convert.ToBoolean(findApi.Execute())) break;
                        replacements++;
                        range.Text = replaceText;
                        var nextStart = Convert.ToInt32(range.End);
                        ComInterop.Release(content);
                        content = doc.Range(nextStart, doc.Content.End);
                    }
                    finally
                    {
                        ComInterop.Release(find);
                    }
                }

                return new { replacements };
            }
            finally
            {
                ComInterop.Release(content);
            }
        });
    }

    public object Save(string? saveAsPath)
    {
        return WithDocument((_, document) =>
        {
            dynamic doc = document;
            if (string.IsNullOrWhiteSpace(saveAsPath))
            {
                doc.Save();
            }
            else
            {
                var fullPath = Path.GetFullPath(saveAsPath);
                Directory.CreateDirectory(Path.GetDirectoryName(fullPath) ?? Environment.CurrentDirectory);
                doc.SaveAs2(fullPath);
                activeDocumentPath = fullPath;
            }

            return new { success = true };
        });
    }

    private object WithDocument(Func<dynamic, object, object> operation)
    {
        using var handle = applications.GetActiveRequired(OrderedProgIds(), "Word 或 WPS 文字未运行，请先打开文档");
        preferredProgId = handle.ProgId;
        dynamic app = handle.Application;
        object? document = ResolveDocument(app);
        try
        {
            if (document is null) throw new OfficeWorkerException("document_not_found", "当前没有活动 Word 文档");
            return operation(app, document);
        }
        finally
        {
            ComInterop.Release(document);
        }
    }

    private object? ResolveDocument(dynamic app)
    {
        if (!string.IsNullOrWhiteSpace(activeDocumentPath))
        {
            object? documents = null;
            try
            {
                documents = app.Documents;
                dynamic documentsApi = documents;
                var count = Convert.ToInt32(documentsApi.Count);
                for (var index = 1; index <= count; index++)
                {
                    object? candidate = documentsApi.Item(index);
                    try
                    {
                        dynamic candidateApi = candidate!;
                        if (string.Equals(Path.GetFullPath(Convert.ToString(candidateApi.FullName) ?? string.Empty), Path.GetFullPath(activeDocumentPath), StringComparison.OrdinalIgnoreCase))
                        {
                            return candidate;
                        }
                    }
                    catch { }
                    ComInterop.Release(candidate);
                }
            }
            finally
            {
                ComInterop.Release(documents);
            }
        }

        return app.ActiveDocument;
    }

    private IEnumerable<string> OrderedProgIds() =>
        preferredProgId is null ? ProgIds : [preferredProgId, .. ProgIds.Where(id => id != preferredProgId)];
}
