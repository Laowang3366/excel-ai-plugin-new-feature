using System.Text.Json;
using System.Text.RegularExpressions;
using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Office;

internal sealed partial class PresentationLinkedContentActionService(OfficeApplicationProvider applications)
{
    private static readonly HashSet<string> Operations = ["inspectLinkedOfficeContent", "refreshLinkedOfficeContent", "relinkLinkedOfficeContent"];

    public static bool Supports(string operation) => Operations.Contains(operation);

    public object Execute(OfficeActionRequest request)
    {
        if (!Supports(request.Operation))
            throw new OfficeWorkerException("unsupported_operation", $"不支持的演示文稿链接操作: {request.Operation}");
        using var context = new PresentationActionContext(applications, request);
        var result = ReadLinks(context, request);
        if (request.Operation != "inspectLinkedOfficeContent") context.Save(request);
        return OfficeActionResults.Done(
            request,
            "com",
            request.Operation == "inspectLinkedOfficeContent" ? "已检查演示文稿链接内容" : "已更新演示文稿链接内容",
            new
            {
                progId = context.ProgId,
                result.Links,
                linkCount = result.Links.Count,
                result.Updated,
                result.Relinked,
                result.Failures,
                manifest = ReadManifest(context.Presentation),
            },
            request.Operation == "inspectLinkedOfficeContent"
                ? Array.Empty<OfficeChange>()
                : [new OfficeChange("linked-content", request.Target, "已更新演示文稿链接内容")]);
    }

    private LinkResult ReadLinks(PresentationActionContext context, OfficeActionRequest request)
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
        object? slides = null;
        try
        {
            slides = context.Presentation.Slides;
            dynamic slidesApi = slides;
            for (var slideIndex = 1; slideIndex <= Convert.ToInt32(slidesApi.Count); slideIndex++)
            {
                object? slide = null;
                object? shapes = null;
                try
                {
                    slide = slidesApi.Item(slideIndex);
                    dynamic slideApi = slide;
                    shapes = slideApi.Shapes;
                    dynamic shapesApi = shapes;
                    for (var shapeIndex = 1; shapeIndex <= Convert.ToInt32(shapesApi.Count); shapeIndex++)
                    {
                        object? shape = null;
                        object? link = null;
                        try
                        {
                            shape = shapesApi.Item(shapeIndex);
                            dynamic shapeApi = shape;
                            try { link = shapeApi.LinkFormat; } catch { continue; }
                            if (link is null) continue;
                            dynamic linkApi = link;
                            var source = Convert.ToString(Safe(() => linkApi.SourceFullName))
                                ?? Convert.ToString(Safe(() => linkApi.SourceName))
                                ?? string.Empty;
                            if (source.Length == 0) continue;
                            var linkId = ReadTag(shape, "WENGGE_LINK_ID");
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
                                    failures.Add(new { slideIndex, name = Convert.ToString(shapeApi.Name), linkId, error = exception.Message });
                                }
                            }
                            else if (request.Operation == "relinkLinkedOfficeContent")
                            {
                                try
                                {
                                    using var sourceContext = OpenLinkedExcelSource(request, newSourcePath);
                                    source = RelinkSource(source, Path.GetFullPath(newSourcePath));
                                    linkApi.SourceFullName = source;
                                    SetTag(shape, "WENGGE_SOURCE_PATH", Path.GetFullPath(newSourcePath));
                                    relinked++;
                                }
                                catch (Exception exception)
                                {
                                    failures.Add(new { slideIndex, name = Convert.ToString(shapeApi.Name), linkId, error = exception.Message });
                                }
                            }
                            var name = Convert.ToString(shapeApi.Name) ?? string.Empty;
                            links.Add(new
                            {
                                version = 1,
                                kind = "shape",
                                slideId = Convert.ToInt32(slideApi.SlideID),
                                slideIndex = Convert.ToInt32(slideApi.SlideIndex),
                                name,
                                shapeName = name,
                                linkId,
                                source,
                                sourceType = ReadTag(shape, "WENGGE_SOURCE_TYPE"),
                                sourceName = ReadTag(shape, "WENGGE_SOURCE_NAME"),
                                range = ReadTag(shape, "WENGGE_SOURCE_RANGE"),
                                managed = ReadTag(shape, "WENGGE_MANAGED").Equals("true", StringComparison.OrdinalIgnoreCase),
                                locator = $"shape:{Convert.ToInt32(slideApi.SlideID)}/{Uri.EscapeDataString(name)}",
                            });
                        }
                        finally { ComInterop.Release(link); ComInterop.Release(shape); }
                    }
                }
                finally { ComInterop.Release(shapes); ComInterop.Release(slide); }
            }
        }
        finally { ComInterop.Release(slides); }
        if (filter.Length > 0 && matched == 0)
            throw new OfficeWorkerException("linked_object_not_found", $"找不到指定 linkId 的演示文稿链接对象: {filter}");
        if (failures.Count > 0 && request.Operation == "refreshLinkedOfficeContent")
            throw new OfficeWorkerException("linked_refresh_failed", "一个或多个演示文稿链接内容刷新失败，未保存本次修改", new { updated, failures });
        if (failures.Count > 0 && request.Operation == "relinkLinkedOfficeContent")
            throw new OfficeWorkerException("linked_relink_failed", "一个或多个演示文稿链接重绑失败，未保存本次修改", new { relinked, failures });
        return new LinkResult(links, updated, relinked, failures);
    }

    private ExcelActionContext? OpenLinkedExcelSource(OfficeActionRequest request, string source)
    {
        var path = ResolveSourcePath(source);
        if (path.Length == 0) throw new OfficeWorkerException("file_not_found", $"找不到 Excel 链接源: {source}");
        return new ExcelActionContext(applications, OfficeHostRouting.SourceRequest(request, "excel", path));
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

    private static object ReadManifest(object presentation)
    {
        var version = ReadTag(presentation, "WENGGE_MANIFEST_VERSION");
        var ids = Array.Empty<string>();
        try { ids = JsonSerializer.Deserialize<string[]>(ReadTag(presentation, "WENGGE_MANAGED_LINK_IDS")) ?? []; } catch { }
        return new { version, managedLinkIds = ids };
    }

    private static string ReadTag(object owner, string name)
    {
        object? tags = null;
        try { tags = ((dynamic)owner).Tags; return Convert.ToString(((dynamic)tags).Item(name)) ?? string.Empty; }
        catch { return string.Empty; }
        finally { ComInterop.Release(tags); }
    }

    private static void SetTag(object owner, string name, string value)
    {
        object? tags = null;
        try
        {
            tags = ((dynamic)owner).Tags;
            try { ((dynamic)tags).Delete(name); } catch { }
            ((dynamic)tags).Add(name, value);
        }
        finally { ComInterop.Release(tags); }
    }

    private static object? Safe(Func<object?> value) { try { return value(); } catch { return null; } }

    private sealed record LinkResult(List<object> Links, int Updated, int Relinked, List<object> Failures);

    [GeneratedRegex("(?i)^(.+?\\.(xlsx|xlsm|xlsb|xls))(?=!|$)")]
    private static partial Regex ExcelSourceRegex();
}
