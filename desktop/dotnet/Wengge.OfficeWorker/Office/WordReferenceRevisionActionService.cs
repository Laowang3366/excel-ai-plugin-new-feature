using System.Text.Json;
using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Office;

internal sealed class WordReferenceRevisionActionService(OfficeApplicationProvider applications, OfficeDocumentService documents)
{
    private static readonly HashSet<string> Operations =
    ["inspectReferences", "manageReferences", "inspectRevisions", "manageRevisions", "compareDocuments", "applyTrackedChanges"];

    public static bool Supports(string operation) => Operations.Contains(operation);

    public object Execute(OfficeActionRequest request)
    {
        var instanceId = request.StringParam("instanceId");
        if (instanceId.Length > 0)
            return documents.WithDocument("word", request.FilePath, instanceId, null, 0, handle =>
            {
                using var borrowedContext = new WordActionContext(handle);
                return ExecuteCore(borrowedContext, request);
            });
        var packageBookmarks = request.Operation == "inspectReferences" && !string.IsNullOrWhiteSpace(request.FilePath)
            ? WordTrackedBookmarkPackageUpdater.Read(Path.GetFullPath(request.FilePath))
            : [];
        Trace($"package-bookmarks:{request.Operation}:{packageBookmarks.Count}");
        object result;
        using (var context = new WordActionContext(applications, request))
            result = ExecuteCore(context, request, packageBookmarks);
        if (request.Operation == "applyTrackedChanges")
        {
            var updates = TrackedBookmarkUpdates(request);
            if (updates.Count > 0)
            {
                var targetPath = string.IsNullOrWhiteSpace(request.OutputPath) ? request.FilePath : request.OutputPath;
                if (string.IsNullOrWhiteSpace(targetPath))
                    throw new OfficeWorkerException("file_required", "修订书签持久化需要 Word 文件路径");
                var savedPath = Path.GetFullPath(targetPath);
                _ = WordTrackedBookmarkPackageUpdater.Normalize(savedPath, updates);
            }
        }
        return result;
    }

    private static IReadOnlyList<WordTrackedBookmarkUpdate> TrackedBookmarkUpdates(OfficeActionRequest request)
    {
        var changes = request.Param("changes");
        if (changes.ValueKind != JsonValueKind.Array) changes = request.Param("edits");
        if (changes.ValueKind != JsonValueKind.Array) return [];
        return changes.EnumerateArray()
            .Where(change => String(change, "command") == "replaceBookmark")
            .Select(change => new WordTrackedBookmarkUpdate(String(change, "name"), String(change, "text")))
            .Where(update => update.Name.Length > 0)
            .ToArray();
    }

    private static object ExecuteCore(
        WordActionContext context,
        OfficeActionRequest request,
        IReadOnlyList<WordPackageBookmark>? packageBookmarks = null)
    {
        if (request.Operation == "inspectReferences")
            return OfficeActionResults.Done(request, "com", "已检查 Word 引用", InspectReferences(context, packageBookmarks ?? []), Array.Empty<OfficeChange>());
        if (request.Operation == "inspectRevisions")
            return OfficeActionResults.Done(request, "com", "已检查 Word 修订", InspectRevisions(context), Array.Empty<OfficeChange>());
        if (request.Operation == "compareDocuments") return Compare(context, request);
        object data;
        if (request.Operation == "manageReferences")
        {
            data = ManageReferences(context, request);
            context.Save(request);
        }
        else if (request.Operation == "manageRevisions")
        {
            data = ManageRevisions(context, request);
            context.Save(request);
        }
        else
        {
            var tracked = ApplyTrackedChanges(context, request);
            context.Save(request);
            var persistedBookmarks = EnsureTrackedBookmarksAfterSave(context.Document, tracked.Replacements);
            data = new
            {
                applied = tracked.Applied,
                tracking = Convert.ToBoolean(context.Document.TrackRevisions),
                restoredBookmarks = tracked.RestoredBookmarks,
                persistedBookmarks,
            };
        }
        return OfficeActionResults.Done(request, "com", "已更新 Word 引用或修订", data,
            [new OfficeChange("word-review", request.Target, "已更新 Word 引用或修订")]);
    }

    private static object InspectReferences(WordActionContext context, IReadOnlyList<WordPackageBookmark>? packageBookmarks = null)
    {
        var bookmarks = new List<object>();
        var bookmarkNames = new HashSet<string>(StringComparer.Ordinal);
        var fields = new List<object>();
        object? bookmarkCollection = null; object? fieldCollection = null;
        try
        {
            bookmarkCollection = context.Document.Bookmarks; dynamic bookmarksApi = bookmarkCollection;
            for (var index = 1; index <= Convert.ToInt32(bookmarksApi.Count); index++)
            {
                object? bookmark = null; object? range = null;
                try
                {
                    bookmark = bookmarksApi.Item(index); dynamic bookmarkApi = bookmark; range = bookmarkApi.Range; dynamic rangeApi = range;
                    var name = Convert.ToString(bookmarkApi.Name) ?? string.Empty;
                    bookmarks.Add(new { name, start = rangeApi.Start, end = rangeApi.End, text = (Convert.ToString(rangeApi.Text) ?? string.Empty).Trim(), source = "com" });
                    if (name.Length > 0) bookmarkNames.Add(name);
                }
                finally { ComInterop.Release(range); ComInterop.Release(bookmark); }
            }
            foreach (var bookmark in (packageBookmarks ?? []).Where(item => !bookmarkNames.Contains(item.Name)))
                bookmarks.Add(new { name = bookmark.Name, start = -1, end = -1, text = bookmark.Text, source = "openxml" });
            Trace($"inspect-bookmarks:com={bookmarkNames.Count}:merged={bookmarks.Count}");
            fieldCollection = context.Document.Fields; dynamic fieldsApi = fieldCollection;
            for (var index = 1; index <= Convert.ToInt32(fieldsApi.Count); index++)
            {
                object? field = null; object? code = null; object? result = null;
                try
                {
                    field = fieldsApi.Item(index); dynamic fieldApi = field; code = fieldApi.Code; result = fieldApi.Result;
                    fields.Add(new { index, type = Safe(() => fieldApi.Type), code = (Convert.ToString(((dynamic)code).Text) ?? string.Empty).Trim(), result = (Convert.ToString(((dynamic)result).Text) ?? string.Empty).Trim() });
                }
                finally { ComInterop.Release(result); ComInterop.Release(code); ComInterop.Release(field); }
            }
            return new
            {
                bookmarks,
                bookmarkCount = bookmarks.Count,
                footnoteCount = Count(() => context.Document.Footnotes.Count),
                endnoteCount = Count(() => context.Document.Endnotes.Count),
                fields,
                fieldCount = fields.Count,
            };
        }
        finally { ComInterop.Release(fieldCollection); ComInterop.Release(bookmarkCollection); }
    }

    private static object ManageReferences(WordActionContext context, OfficeActionRequest request)
    {
        var command = request.StringParam("command", "updateFields");
        object? range = ResolveRange(context, request);
        try
        {
            switch (command)
            {
                case "createBookmark":
                case "addBookmark":
                    var name = request.StringParam("name");
                    object? bookmarks = null;
                    try { bookmarks = context.Document.Bookmarks; ((dynamic)bookmarks).Add(name, range); }
                    finally { ComInterop.Release(bookmarks); }
                    break;
                case "deleteBookmark":
                    object? collection = null; object? bookmark = null;
                    try { collection = context.Document.Bookmarks; bookmark = ((dynamic)collection).Item(request.StringParam("name")); ((dynamic)bookmark).Delete(); }
                    finally { ComInterop.Release(bookmark); ComInterop.Release(collection); }
                    break;
                case "addFootnote":
                    object? footnotes = null; object? footnote = null;
                    try { footnotes = context.Document.Footnotes; footnote = ((dynamic)footnotes).Add(range, Type.Missing, request.StringParam("text")); }
                    finally { ComInterop.Release(footnote); ComInterop.Release(footnotes); }
                    break;
                case "addEndnote":
                    object? endnotes = null; object? endnote = null;
                    try { endnotes = context.Document.Endnotes; endnote = ((dynamic)endnotes).Add(range, Type.Missing, request.StringParam("text")); }
                    finally { ComInterop.Release(endnote); ComInterop.Release(endnotes); }
                    break;
                case "addCaption":
                    dynamic rangeApi = range!;
                    var label = request.StringParam("label", "Figure");
                    try
                    {
                        rangeApi.InsertCaption(label, request.StringParam("title"), Type.Missing, 0, false);
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        object? labels = null; object? createdLabel = null;
                        try
                        {
                            labels = context.App.CaptionLabels;
                            createdLabel = ((dynamic)labels).Add(label);
                            rangeApi.InsertCaption(label, request.StringParam("title"), Type.Missing, 0, false);
                        }
                        finally { ComInterop.Release(createdLabel); ComInterop.Release(labels); }
                    }
                    break;
                case "addCrossReference":
                    rangeApi = range!;
                    if (request.StringParam("referenceType", "bookmark").Equals("bookmark", StringComparison.OrdinalIgnoreCase))
                    {
                        object? crossReferenceFields = null; object? crossReferenceField = null;
                        try
                        {
                            crossReferenceFields = context.Document.Fields;
                            crossReferenceField = ((dynamic)crossReferenceFields).Add(range, 3, request.StringParam("item") + " \\h", true);
                        }
                        finally { ComInterop.Release(crossReferenceField); ComInterop.Release(crossReferenceFields); }
                    }
                    else
                    {
                        rangeApi.InsertCrossReference(request.StringParam("referenceType"), -1, request.StringParam("item"), true, false, false, " ");
                    }
                    break;
                case "addTableOfFigures":
                    object? figures = null; object? table = null;
                    try
                    {
                        figures = context.Document.TablesOfFigures;
                        table = ((dynamic)figures).Add(range, request.StringParam("label", "Figure"), true, false, 1, 9, false, "", true, true, "", true, true);
                    }
                    finally { ComInterop.Release(table); ComInterop.Release(figures); }
                    break;
                case "updateFields":
                    object? fields = null;
                    try { fields = context.Document.Fields; ((dynamic)fields).Update(); }
                    finally { ComInterop.Release(fields); }
                    object? tables = null;
                    try
                    {
                        tables = context.Document.TablesOfContents; dynamic tablesApi = tables;
                        for (var index = 1; index <= Convert.ToInt32(tablesApi.Count); index++) { object? toc = null; try { toc = tablesApi.Item(index); ((dynamic)toc).Update(); } finally { ComInterop.Release(toc); } }
                    }
                    finally { ComInterop.Release(tables); }
                    break;
                default: throw new OfficeWorkerException("unsupported_operation", $"不支持的 Word 引用命令: {command}");
            }
            return new { command, references = InspectReferences(context) };
        }
        finally { ComInterop.Release(range); }
    }

    private static object InspectRevisions(WordActionContext context)
    {
        var revisions = new List<object>();
        object? collection = null;
        try
        {
            collection = context.Document.Revisions; dynamic revisionsApi = collection;
            for (var index = 1; index <= Convert.ToInt32(revisionsApi.Count); index++)
            {
                object? revision = null; object? range = null;
                try
                {
                    revision = revisionsApi.Item(index); dynamic revisionApi = revision; range = revisionApi.Range; dynamic rangeApi = range;
                    revisions.Add(new { index, author = Safe(() => revisionApi.Author), type = Safe(() => revisionApi.Type), date = Safe(() => revisionApi.Date), text = (Convert.ToString(rangeApi.Text) ?? string.Empty).Trim(), start = rangeApi.Start, end = rangeApi.End });
                }
                finally { ComInterop.Release(range); ComInterop.Release(revision); }
            }
            return new { revisions, revisionCount = revisions.Count, trackRevisions = Convert.ToBoolean(context.Document.TrackRevisions) };
        }
        finally { ComInterop.Release(collection); }
    }

    private static object ManageRevisions(WordActionContext context, OfficeActionRequest request)
    {
        var command = request.StringParam("command", "acceptAll");
        object? revisions = null;
        var changed = 0;
        try
        {
            revisions = context.Document.Revisions; dynamic revisionsApi = revisions;
            if (command == "acceptAll") { changed = Convert.ToInt32(revisionsApi.Count); revisionsApi.AcceptAll(); }
            else if (command == "rejectAll") { changed = Convert.ToInt32(revisionsApi.Count); revisionsApi.RejectAll(); }
            else if (command is "accept" or "reject")
            {
                var author = request.StringParam("author");
                var type = request.IntParam("revisionType", int.MinValue);
                for (var index = Convert.ToInt32(revisionsApi.Count); index >= 1; index--)
                {
                    object? revision = null;
                    try
                    {
                        revision = revisionsApi.Item(index); dynamic revisionApi = revision;
                        if (author.Length > 0 && !string.Equals(Convert.ToString(revisionApi.Author), author, StringComparison.OrdinalIgnoreCase)) continue;
                        if (type != int.MinValue && Convert.ToInt32(revisionApi.Type) != type) continue;
                        if (command == "accept") revisionApi.Accept(); else revisionApi.Reject();
                        changed++;
                    }
                    finally { ComInterop.Release(revision); }
                }
            }
            else if (command == "track") context.Document.TrackRevisions = request.BoolParam("enabled", true);
            else throw new OfficeWorkerException("unsupported_operation", $"不支持的修订命令: {command}");
            return new { command, changed, remaining = Convert.ToInt32(context.Document.Revisions.Count) };
        }
        finally { ComInterop.Release(revisions); }
    }

    private static TrackedChangesApplication ApplyTrackedChanges(WordActionContext context, OfficeActionRequest request)
    {
        var previous = Convert.ToBoolean(context.Document.TrackRevisions);
        context.Document.TrackRevisions = true;
        var applied = 0;
        var replacedBookmarks = new List<TrackedBookmarkReplacement>();
        try
        {
            var changes = request.Param("changes");
            if (changes.ValueKind != JsonValueKind.Array) changes = request.Param("edits");
            if (changes.ValueKind == JsonValueKind.Array)
            {
                foreach (var change in changes.EnumerateArray())
                {
                    var command = String(change, "command");
                    if (command == "insert")
                    {
                        var text = String(change, "text");
                        var position = String(change, "position") switch
                        {
                            "start" => 0,
                            "end" or "" => Math.Max(0, Convert.ToInt32(context.Document.Content.End) - 1),
                            _ => change.TryGetProperty("position", out var positionValue) && positionValue.TryGetInt32(out var explicitPosition)
                                ? Math.Clamp(explicitPosition, 0, Math.Max(0, Convert.ToInt32(context.Document.Content.End) - 1))
                                : Math.Max(0, Convert.ToInt32(context.Document.Content.End) - 1),
                        };
                        object? insertion = null;
                        try { insertion = context.Document.Range(position, position); ((dynamic)insertion).InsertAfter(text); applied++; }
                        finally { ComInterop.Release(insertion); }
                        continue;
                    }
                    if (command == "replaceBookmark")
                    {
                        var name = String(change, "name");
                        var text = String(change, "text");
                        object? bookmarks = null;
                        object? bookmark = null;
                        object? target = null;
                        try
                        {
                            bookmarks = context.Document.Bookmarks;
                            if (name.Length == 0 || !Convert.ToBoolean(((dynamic)bookmarks).Exists(name)))
                                throw new OfficeWorkerException("bookmark_not_found", $"replaceBookmark 找不到书签: {name}");
                            bookmark = ((dynamic)bookmarks).Item(name);
                            target = ((dynamic)bookmark).Range;
                            var start = Convert.ToInt32(((dynamic)target).Start);
                            ((dynamic)target).Text = text;
                            replacedBookmarks.Add(new TrackedBookmarkReplacement(name, start, text));
                            applied++;
                        }
                        finally
                        {
                            ComInterop.Release(target);
                            ComInterop.Release(bookmark);
                            ComInterop.Release(bookmarks);
                        }
                        continue;
                    }
                    if (command == "replaceContentControl")
                    {
                        var tag = String(change, "tag");
                        var title = String(change, "title");
                        var text = String(change, "text");
                        object? controls = null;
                        try
                        {
                            controls = context.Document.ContentControls;
                            dynamic controlsApi = controls;
                            for (var index = 1; index <= Convert.ToInt32(controlsApi.Count); index++)
                            {
                                object? control = null;
                                object? controlRange = null;
                                try
                                {
                                    control = controlsApi.Item(index);
                                    dynamic controlApi = control;
                                    if (tag.Length > 0 && !string.Equals(Convert.ToString(controlApi.Tag), tag, StringComparison.Ordinal)
                                        && title.Length > 0 && !string.Equals(Convert.ToString(controlApi.Title), title, StringComparison.Ordinal)) continue;
                                    if (tag.Length > 0 && title.Length == 0 && !string.Equals(Convert.ToString(controlApi.Tag), tag, StringComparison.Ordinal)) continue;
                                    if (title.Length > 0 && tag.Length == 0 && !string.Equals(Convert.ToString(controlApi.Title), title, StringComparison.Ordinal)) continue;
                                    try { controlApi.LockContents = false; } catch { }
                                    controlRange = controlApi.Range;
                                    ((dynamic)controlRange).Text = text;
                                    applied++;
                                }
                                finally { ComInterop.Release(controlRange); ComInterop.Release(control); }
                            }
                        }
                        finally { ComInterop.Release(controls); }
                        continue;
                    }
                    var find = String(change, "find"); var replace = String(change, "replace");
                    if (find.Length == 0) continue;
                    object? range = null; object? finder = null;
                    object? replacement = null;
                    try
                    {
                        range = context.Document.Content; dynamic rangeApi = range; finder = rangeApi.Find; dynamic findApi = finder;
                        var matchCase = Bool(change, "matchCase");
                        var replaceAll = !change.TryGetProperty("all", out var allProperty) || allProperty.ValueKind != JsonValueKind.False;
                        var sourceText = Convert.ToString(rangeApi.Text) ?? string.Empty;
                        var matches = CountOccurrences(sourceText, find, matchCase);
                        findApi.ClearFormatting();
                        replacement = findApi.Replacement; ((dynamic)replacement).ClearFormatting(); ((dynamic)replacement).Text = replace;
                        findApi.Text = find; findApi.MatchCase = matchCase; findApi.Forward = true; findApi.Wrap = 1;
                        _ = findApi.Execute(Replace: replaceAll ? 2 : 1);
                        applied += replaceAll ? matches : Math.Min(matches, 1);
                    }
                    finally { ComInterop.Release(replacement); ComInterop.Release(finder); ComInterop.Release(range); }
                }
            }
        }
        finally
        {
            if (!request.BoolParam("keepTracking") && request.BoolParam("restoreTracking", true))
                context.Document.TrackRevisions = previous;
        }
        var restoredBookmarks = RestoreTrackedBookmarks(context.Document, replacedBookmarks);
        return new TrackedChangesApplication(applied, restoredBookmarks, replacedBookmarks);
    }

    private static IReadOnlyList<string> EnsureTrackedBookmarksAfterSave(
        dynamic document,
        IReadOnlyList<TrackedBookmarkReplacement> replacements)
    {
        if (replacements.Count == 0) return [];
        if (replacements.All(replacement => BookmarkContains(document, replacement.Name, replacement.Text)))
            return replacements.Select(replacement => replacement.Name).ToArray();

        _ = RestoreTrackedBookmarks(document, replacements);
        document.Save();
        foreach (var replacement in replacements)
        {
            if (!BookmarkContains(document, replacement.Name, replacement.Text))
                throw new OfficeWorkerException(
                    "bookmark_restore_failed",
                    $"保存后书签恢复校验失败: {replacement.Name}");
        }
        return replacements.Select(replacement => replacement.Name).ToArray();
    }

    private static bool BookmarkContains(dynamic document, string name, string expected)
    {
        object? bookmarks = null;
        object? bookmark = null;
        object? range = null;
        try
        {
            bookmarks = document.Bookmarks;
            dynamic bookmarksApi = bookmarks;
            if (!Convert.ToBoolean(bookmarksApi.Exists(name))) return false;
            bookmark = bookmarksApi.Item(name);
            range = ((dynamic)bookmark).Range;
            return (Convert.ToString(((dynamic)range).Text) ?? string.Empty).Contains(expected, StringComparison.Ordinal);
        }
        catch { return false; }
        finally
        {
            ComInterop.Release(range);
            ComInterop.Release(bookmark);
            ComInterop.Release(bookmarks);
        }
    }

    private static IReadOnlyList<string> RestoreTrackedBookmarks(dynamic document, IReadOnlyList<TrackedBookmarkReplacement> replacements)
    {
        if (replacements.Count == 0) return [];
        var tracking = Convert.ToBoolean(document.TrackRevisions);
        object? bookmarks = null;
        object? content = null;
        var restored = new List<string>();
        try
        {
            document.TrackRevisions = false;
            bookmarks = document.Bookmarks;
            dynamic bookmarksApi = bookmarks;
            content = document.Content;
            var contentText = Convert.ToString(((dynamic)content).Text) ?? string.Empty;
            var contentEnd = Math.Max(0, Convert.ToInt32(((dynamic)content).End) - 1);
            foreach (var replacement in replacements)
            {
                if (Convert.ToBoolean(bookmarksApi.Exists(replacement.Name)))
                {
                    object? existing = null;
                    try { existing = bookmarksApi.Item(replacement.Name); ((dynamic)existing).Delete(); }
                    finally { ComInterop.Release(existing); }
                }
                var start = Math.Clamp(replacement.Start, 0, contentEnd);
                var end = Math.Clamp(start + replacement.Text.Length, start, contentEnd);
                if (!WordRangeMatches(document, start, end, replacement.Text))
                {
                    var located = ClosestText(contentText, replacement.Text, replacement.Start);
                    if (located < 0)
                        throw new OfficeWorkerException("bookmark_restore_failed", $"修订后找不到书签文本: {replacement.Name}");
                    start = located;
                    end = Math.Min(contentEnd, located + replacement.Text.Length);
                }
                object? range = null;
                object? bookmark = null;
                try
                {
                    range = document.Range(start, end);
                    bookmark = bookmarksApi.Add(replacement.Name, range);
                }
                finally
                {
                    ComInterop.Release(bookmark);
                    ComInterop.Release(range);
                }
                if (!Convert.ToBoolean(bookmarksApi.Exists(replacement.Name)))
                    throw new OfficeWorkerException("bookmark_restore_failed", $"修订后书签恢复校验失败: {replacement.Name}");
                restored.Add(replacement.Name);
            }
            return restored;
        }
        finally
        {
            ComInterop.Release(content);
            ComInterop.Release(bookmarks);
            document.TrackRevisions = tracking;
        }
    }

    private static bool WordRangeMatches(dynamic document, int start, int end, string expected)
    {
        object? range = null;
        try
        {
            range = document.Range(start, end);
            return string.Equals(Convert.ToString(((dynamic)range).Text) ?? string.Empty, expected, StringComparison.Ordinal);
        }
        catch { return false; }
        finally { ComInterop.Release(range); }
    }

    private static int ClosestText(string source, string value, int preferredIndex)
    {
        var closest = -1;
        var distance = int.MaxValue;
        for (var index = source.IndexOf(value, StringComparison.Ordinal); index >= 0; index = source.IndexOf(value, index + 1, StringComparison.Ordinal))
        {
            var candidateDistance = Math.Abs(index - preferredIndex);
            if (candidateDistance >= distance) continue;
            closest = index;
            distance = candidateDistance;
        }
        return closest;
    }

    private sealed record TrackedBookmarkReplacement(string Name, int Start, string Text);
    private sealed record TrackedChangesApplication(
        int Applied,
        IReadOnlyList<string> RestoredBookmarks,
        IReadOnlyList<TrackedBookmarkReplacement> Replacements);

    private static object Compare(WordActionContext context, OfficeActionRequest request)
    {
        var comparePath = Path.GetFullPath(request.StringParam("comparePath", request.StringParam("revisedFilePath")));
        if (!File.Exists(comparePath)) throw new OfficeWorkerException("file_not_found", $"对比文档不存在: {comparePath}");
        var output = request.OutputPath ?? Path.Combine(Path.GetDirectoryName(request.FilePath) ?? Environment.CurrentDirectory, "comparison.docx");
        output = Path.GetFullPath(output);
        object? documents = null; object? revised = null; object? comparison = null;
        try
        {
            documents = context.App.Documents; revised = ((dynamic)documents).Open(comparePath, false, true);
            comparison = context.App.CompareDocuments(context.Document, revised, 2, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, "Wengge", true);
            ((dynamic)comparison).SaveAs2(output);
            return OfficeActionResults.Done(request, "com", "已生成文档对比结果", new { outputPath = output, revisionCount = Safe(() => ((dynamic)comparison).Revisions.Count) }, [new OfficeChange("document-comparison", output, "已生成文档对比结果")], output);
        }
        finally { if (comparison is not null) { try { ((dynamic)comparison).Close(0); } catch { } } if (revised is not null) { try { ((dynamic)revised).Close(0); } catch { } } ComInterop.Release(comparison); ComInterop.Release(revised); ComInterop.Release(documents); }
    }

    private static object? ResolveRange(WordActionContext context, OfficeActionRequest request)
    {
        var bookmark = request.StringParam("bookmark");
        if (bookmark.Length > 0)
        {
            object? bookmarks = null; object? item = null;
            try { bookmarks = context.Document.Bookmarks; item = ((dynamic)bookmarks).Item(bookmark); return ((dynamic)item).Range; }
            finally { ComInterop.Release(item); ComInterop.Release(bookmarks); }
        }
        return context.Document.Range(Math.Max(0, request.IntParam("start", Convert.ToInt32(context.Document.Content.End) - 1)), Math.Max(0, request.IntParam("end", Convert.ToInt32(context.Document.Content.End) - 1)));
    }

    private static string String(JsonElement value, string name) => value.TryGetProperty(name, out var property) && property.ValueKind == JsonValueKind.String ? property.GetString() ?? string.Empty : string.Empty;
    private static bool Bool(JsonElement value, string name) => value.TryGetProperty(name, out var property) && property.ValueKind is JsonValueKind.True or JsonValueKind.False && property.GetBoolean();
    private static void Trace(string message)
    {
        if (Environment.GetEnvironmentVariable("WENGGE_OFFICE_SMOKE") == "1")
            Console.Error.WriteLine($"[office-smoke] word:{message}");
    }
    private static int CountOccurrences(string source, string value, bool matchCase)
    {
        var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var count = 0;
        for (var start = 0; start < source.Length; count++)
        {
            var index = source.IndexOf(value, start, comparison);
            if (index < 0) break;
            start = index + value.Length;
        }
        return count;
    }
    private static int Count(Func<object> value) { try { return Convert.ToInt32(value()); } catch { return 0; } }
    private static object? Safe(Func<object?> value) { try { return value(); } catch { return null; } }
}
