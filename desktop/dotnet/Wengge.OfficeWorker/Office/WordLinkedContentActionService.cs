using System.Text.Json;
using System.Text.RegularExpressions;
using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Office;

internal sealed partial class WordLinkedContentActionService(OfficeApplicationProvider applications)
{
    private static readonly HashSet<string> Operations = ["inspectLinkedOfficeContent", "refreshLinkedOfficeContent", "relinkLinkedOfficeContent"];

    public static bool Supports(string operation) => Operations.Contains(operation);

    public object Execute(OfficeActionRequest request)
    {
        if (!Supports(request.Operation)) throw new OfficeWorkerException("unsupported_operation", $"不支持的 Word 链接操作: {request.Operation}");
        using var context = new WordActionContext(applications, request);
        var result = ReadLinks(context, request);
        if (request.Operation != "inspectLinkedOfficeContent") context.Save(request);
        return OfficeActionResults.Done(
            request,
            "com",
            request.Operation == "inspectLinkedOfficeContent" ? "已检查 Word 链接内容" : "已更新 Word 链接内容",
            new
            {
                result.Links,
                linkCount = result.Links.Count,
                result.Updated,
                result.Relinked,
                result.Failures,
                manifest = ReadManifest(context.Document),
            },
            request.Operation == "inspectLinkedOfficeContent"
                ? Array.Empty<OfficeChange>()
                : [new OfficeChange("linked-content", request.Target, "已更新 Word 链接内容")]);
    }

    private LinkResult ReadLinks(WordActionContext context, OfficeActionRequest request)
    {
        var links = new List<object>();
        var failures = new List<object>();
        var updated = 0;
        var relinked = 0;
        var matched = 0;
        var filter = request.StringParam("linkId");
        var newSourcePath = request.StringParam("sourcePath", request.StringParam("newSourcePath"));
        if (request.Operation == "relinkLinkedOfficeContent" && (filter.Length == 0 || newSourcePath.Length == 0))
            throw new OfficeWorkerException("invalid_params", "重绑链接需要 params.linkId 和 params.sourcePath");
        object? inlineShapes = null;
        object? shapes = null;
        var managedIds = ReadManagedIds(context.Document);
        try
        {
            inlineShapes = context.Document.InlineShapes;
            ReadCollection(inlineShapes, "inlineShape", context, request, filter, newSourcePath, managedIds, links, failures, ref updated, ref relinked, ref matched);
            shapes = context.Document.Shapes;
            ReadCollection(shapes, "shape", context, request, filter, newSourcePath, [], links, failures, ref updated, ref relinked, ref matched);
        }
        finally { ComInterop.Release(shapes); ComInterop.Release(inlineShapes); }
        if (filter.Length > 0 && matched == 0)
            throw new OfficeWorkerException("linked_object_not_found", $"找不到指定 linkId 的 Word 链接对象: {filter}");
        if (failures.Count > 0 && request.Operation == "refreshLinkedOfficeContent")
            throw new OfficeWorkerException("linked_refresh_failed", "一个或多个 Word 链接内容刷新失败，未保存本次修改", new { updated, failures });
        if (failures.Count > 0 && request.Operation == "relinkLinkedOfficeContent")
            throw new OfficeWorkerException("linked_relink_failed", "一个或多个 Word 链接重绑失败，未保存本次修改", new { relinked, failures });
        return new LinkResult(links, updated, relinked, failures);
    }

    private void ReadCollection(
        object collection,
        string kind,
        WordActionContext context,
        OfficeActionRequest request,
        string filter,
        string newSourcePath,
        string[] managedIds,
        List<object> output,
        List<object> failures,
        ref int updated,
        ref int relinked,
        ref int matched)
    {
        dynamic collectionApi = collection;
        for (var index = 1; index <= Convert.ToInt32(collectionApi.Count); index++)
        {
            object? shape = null;
            object? link = null;
            try
            {
                shape = collectionApi.Item(index);
                dynamic shapeApi = shape;
                try { link = shapeApi.LinkFormat; } catch { continue; }
                if (link is null) continue;
                dynamic linkApi = link;
                var source = Convert.ToString(Safe(() => linkApi.SourceFullName))
                    ?? Convert.ToString(Safe(() => linkApi.SourceName))
                    ?? string.Empty;
                if (source.Length == 0) continue;
                var metadata = ReadMetadata(shape);
                var linkId = metadata.TryGetValue("linkId", out var id) ? Convert.ToString(id) ?? string.Empty : string.Empty;
                if (linkId.Length == 0 && index <= managedIds.Length) linkId = managedIds[index - 1];
                if (filter.Length > 0 && !linkId.Equals(filter, StringComparison.Ordinal)) continue;
                matched++;
                if (request.Operation == "refreshLinkedOfficeContent")
                {
                    try
                    {
                        using var sourceContext = OpenLinkedExcelSource(request, source);
                        linkApi.Update();
                        updated++;
                    }
                    catch (Exception exception)
                    {
                        failures.Add(new { kind, index, linkId, error = exception.Message });
                    }
                }
                else if (request.Operation == "relinkLinkedOfficeContent")
                {
                    try
                    {
                        using var sourceContext = OpenLinkedExcelSource(request, newSourcePath);
                        source = RelinkSource(source, Path.GetFullPath(newSourcePath));
                        try { linkApi.SourceFullName = source; } catch { linkApi.SourceName = source; }
                        metadata["source"] = Path.GetFullPath(newSourcePath);
                        try { shapeApi.AlternativeText = "WENGGE_MANIFEST:" + JsonSerializer.Serialize(metadata); } catch { }
                        relinked++;
                    }
                    catch (Exception exception)
                    {
                        failures.Add(new { kind, index, linkId, error = exception.Message });
                    }
                }
                output.Add(new
                {
                    version = 1,
                    kind,
                    index,
                    name = Safe(() => shapeApi.Name),
                    linkId,
                    source,
                    sourceType = metadata.GetValueOrDefault("sourceType"),
                    sourceName = metadata.GetValueOrDefault("sourceName"),
                    range = metadata.GetValueOrDefault("range"),
                    managed = metadata.Count > 0,
                    locator = $"{kind}:{index}",
                    metadata,
                });
            }
            finally { ComInterop.Release(link); ComInterop.Release(shape); }
        }
    }

    private ExcelActionContext? OpenLinkedExcelSource(OfficeActionRequest request, string source)
    {
        var path = ResolveSourcePath(source);
        if (path.Length == 0) throw new OfficeWorkerException("file_not_found", $"找不到 Excel 链接源: {source}");
        return new ExcelActionContext(applications, OfficeHostRouting.SourceRequest(request, "excel", path));
    }

    private static Dictionary<string, object?> ReadMetadata(object shape)
    {
        string text;
        try { text = Convert.ToString(((dynamic)shape).AlternativeText) ?? string.Empty; } catch { return []; }
        if (!text.StartsWith("WENGGE_MANIFEST:", StringComparison.Ordinal)) return [];
        try { return JsonSerializer.Deserialize<Dictionary<string, object?>>(text[16..]) ?? []; }
        catch { return []; }
    }

    private static string ResolveSourcePath(string source)
    {
        if (File.Exists(source)) return Path.GetFullPath(source);
        var match = ExcelSourceRegex().Match(source);
        return match.Success && File.Exists(match.Groups[1].Value) ? Path.GetFullPath(match.Groups[1].Value) : string.Empty;
    }

    private static string RelinkSource(string oldSource, string newPath)
    {
        var separator = oldSource.IndexOf('!');
        if (separator < 0) return newPath;
        var suffix = oldSource[separator..];
        var oldName = Path.GetFileName(oldSource[..separator]);
        var newName = Path.GetFileName(newPath);
        if (oldName.Length > 0 && newName.Length > 0) suffix = suffix.Replace($"[{oldName}]", $"[{newName}]", StringComparison.OrdinalIgnoreCase);
        return newPath + suffix;
    }

    private static object ReadManifest(object document)
    {
        object? variables = null;
        object? versionVariable = null;
        object? idsVariable = null;
        var version = string.Empty;
        var ids = Array.Empty<string>();
        try
        {
            variables = ((dynamic)document).Variables;
            try { versionVariable = ((dynamic)variables).Item("WENGGE_MANIFEST_VERSION"); version = Convert.ToString(((dynamic)versionVariable).Value) ?? string.Empty; } catch { }
            try
            {
                idsVariable = ((dynamic)variables).Item("WENGGE_MANAGED_LINK_IDS");
                var serializedIds = Convert.ToString((object?)((dynamic)idsVariable).Value) ?? string.Empty;
                ids = JsonSerializer.Deserialize<string[]>(serializedIds) ?? Array.Empty<string>();
            }
            catch { }
            return new { version, managedLinkIds = ids };
        }
        finally
        {
            ComInterop.Release(idsVariable);
            ComInterop.Release(versionVariable);
            ComInterop.Release(variables);
        }
    }

    private static string[] ReadManagedIds(object document)
    {
        object? variables = null;
        object? item = null;
        try
        {
            variables = ((dynamic)document).Variables;
            item = ((dynamic)variables).Item("WENGGE_MANAGED_LINK_IDS");
            var serialized = Convert.ToString((object?)((dynamic)item).Value) ?? string.Empty;
            return JsonSerializer.Deserialize<string[]>(serialized) ?? [];
        }
        catch { return []; }
        finally { ComInterop.Release(item); ComInterop.Release(variables); }
    }

    private static object? Safe(Func<object?> value) { try { return value(); } catch { return null; } }

    private sealed record LinkResult(List<object> Links, int Updated, int Relinked, List<object> Failures);

    [GeneratedRegex("(?i)^(.+?\\.(xlsx|xlsm|xlsb|xls))(?=!|$)")]
    private static partial Regex ExcelSourceRegex();
}
