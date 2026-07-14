using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Office;

internal sealed record WordTrackedBookmarkUpdate(string Name, string Text);
internal sealed record WordPackageBookmark(string Name, string Text);

internal static class WordTrackedBookmarkPackageUpdater
{
    public static IReadOnlyList<WordPackageBookmark> Read(string filePath)
    {
        using var package = WordprocessingDocument.Open(filePath, false);
        var document = package.MainDocumentPart?.Document;
        if (document is null) return [];
        var elements = document.Descendants().ToList();
        var bookmarks = new List<WordPackageBookmark>();
        foreach (var start in elements.OfType<BookmarkStart>())
        {
            var name = start.Name?.Value ?? string.Empty;
            var id = start.Id?.Value;
            if (name.Length == 0 || id is null) continue;
            var startIndex = elements.IndexOf(start);
            var endIndex = elements.FindIndex(startIndex + 1, element =>
                element is BookmarkEnd end && string.Equals(end.Id?.Value, id, StringComparison.Ordinal));
            if (endIndex < 0) continue;
            var text = string.Concat(elements.Skip(startIndex + 1).Take(endIndex - startIndex - 1)
                .OfType<Text>().Select(item => item.Text)).Replace("\u200B", string.Empty, StringComparison.Ordinal);
            bookmarks.Add(new WordPackageBookmark(name, text.Trim()));
        }
        return bookmarks;
    }

    public static IReadOnlyList<string> Normalize(string filePath, IReadOnlyList<WordTrackedBookmarkUpdate> updates)
    {
        if (updates.Count == 0) return [];
        using var package = WordprocessingDocument.Open(filePath, true);
        var document = package.MainDocumentPart?.Document
            ?? throw new OfficeWorkerException("invalid_document", "Word 文档缺少主文档部件");
        var normalized = new List<string>();
        foreach (var update in updates)
        {
            var start = document.Descendants<BookmarkStart>()
                .FirstOrDefault(item => string.Equals(item.Name?.Value, update.Name, StringComparison.Ordinal));
            if (start is null)
                throw new OfficeWorkerException("bookmark_restore_failed", $"落盘后找不到书签起点: {update.Name}");
            var id = start.Id?.Value;
            var end = document.Descendants<BookmarkEnd>()
                .FirstOrDefault(item => string.Equals(item.Id?.Value, id, StringComparison.Ordinal));
            if (end is null)
                throw new OfficeWorkerException("bookmark_restore_failed", $"落盘后找不到书签终点: {update.Name}");

            var insertion = FindEnclosedInsertion(start, end, update.Text)
                ?? document.Descendants<InsertedRun>().FirstOrDefault(item => ContainsText(item, update.Text));
            if (insertion is null)
                throw new OfficeWorkerException("bookmark_restore_failed", $"落盘后找不到书签修订文本: {update.Name}");
            AnchorAcrossStableRun(start, end, insertion);
            normalized.Add(update.Name);
        }
        document.Save();
        return normalized;
    }

    private static InsertedRun? FindEnclosedInsertion(BookmarkStart start, BookmarkEnd end, string text)
    {
        if (start.Parent is null || !ReferenceEquals(start.Parent, end.Parent)) return null;
        var siblings = start.Parent.ChildElements;
        var startIndex = -1;
        var endIndex = -1;
        for (var index = 0; index < siblings.Count; index++)
        {
            if (ReferenceEquals(siblings[index], start)) startIndex = index;
            if (ReferenceEquals(siblings[index], end)) endIndex = index;
        }
        if (startIndex < 0 || endIndex <= startIndex) return null;
        for (var index = startIndex + 1; index < endIndex; index++)
        {
            if (siblings[index] is InsertedRun insertion && ContainsText(insertion, text)) return insertion;
        }
        return null;
    }

    private static bool ContainsText(OpenXmlElement element, string expected) =>
        string.Concat(element.Descendants<Text>().Select(item => item.Text)).Contains(expected, StringComparison.Ordinal);

    private static void AnchorAcrossStableRun(BookmarkStart start, BookmarkEnd end, InsertedRun insertion)
    {
        var paragraph = insertion.Parent
            ?? throw new OfficeWorkerException("bookmark_restore_failed", "修订文本缺少父级段落");
        var insertionIndex = -1;
        OpenXmlElement? anchor = null;
        for (var index = 0; index < paragraph.ChildElements.Count; index++)
        {
            var child = paragraph.ChildElements[index];
            if (ReferenceEquals(child, insertion)) insertionIndex = index;
            if (insertionIndex < 0 && child is Run run
                && run.Descendants<Text>().Any(item => item.Text.Length > 0))
                anchor = child;
        }
        if (insertionIndex < 0)
            throw new OfficeWorkerException("bookmark_restore_failed", "无法定位修订文本节点");

        start.Remove();
        end.Remove();
        if (anchor is null)
        {
            anchor = new Run(new Text("\u200B"));
            paragraph.InsertBefore(anchor, insertion);
        }
        paragraph.InsertBefore(start, anchor);
        paragraph.InsertAfter(end, insertion);
    }
}
