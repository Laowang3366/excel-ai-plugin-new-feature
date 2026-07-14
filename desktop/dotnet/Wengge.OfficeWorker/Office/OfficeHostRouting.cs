using System.Text.Json;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Office;

internal static class OfficeHostRouting
{
    public static string[] ProgIds(string app, string? host) => (app, (host ?? string.Empty).Trim().ToLowerInvariant()) switch
    {
        ("excel", "wps" or "ket") => ["Ket.Application"],
        ("excel", "excel" or "microsoft" or "office") => ["Excel.Application"],
        ("excel", _) => ["Excel.Application", "Ket.Application"],
        ("word", "wps" or "kwps") => ["Kwps.Application", "Wps.Application"],
        ("word", "word" or "microsoft" or "office") => ["Word.Application"],
        ("word", _) => ["Word.Application", "Kwps.Application", "Wps.Application"],
        ("presentation", "wps" or "wpp") => ["Wpp.Application", "Kwpp.Application"],
        ("presentation", "powerpoint" or "ppt" or "microsoft" or "office") => ["PowerPoint.Application"],
        ("presentation", _) => ["PowerPoint.Application", "Wpp.Application", "Kwpp.Application"],
        _ => throw new OfficeWorkerException("unsupported_app", $"不支持的 Office 应用: {app}"),
    };

    public static void Validate(string app, string? requestedHost, string actualProgId)
    {
        if (string.IsNullOrWhiteSpace(requestedHost)) return;
        if (!ProgIds(app, requestedHost).Contains(actualProgId, StringComparer.OrdinalIgnoreCase))
            throw new OfficeWorkerException("office_host_mismatch", $"instanceId 对应 {actualProgId}，与请求宿主 {requestedHost} 不一致");
    }

    public static bool IsWps(string progId) => progId.Contains("wps", StringComparison.OrdinalIgnoreCase)
        || progId.Contains("ket", StringComparison.OrdinalIgnoreCase)
        || progId.Contains("wpp", StringComparison.OrdinalIgnoreCase);

    public static bool RequestsWps(string? host) => (host ?? string.Empty).Trim().ToLowerInvariant() is "wps" or "ket" or "kwps" or "wpp";

    public static OfficeActionRequest SourceRequest(OfficeActionRequest request, string app, string? filePath)
    {
        var parameters = request.Params.ValueKind == JsonValueKind.Object
            ? request.Params.EnumerateObject().ToDictionary(property => property.Name, property => property.Value.Clone(), StringComparer.Ordinal)
            : new Dictionary<string, JsonElement>(StringComparer.Ordinal);
        parameters.Remove("host");
        parameters.Remove("instanceId");
        if (request.StringParam("sourceHost") is { Length: > 0 } sourceHost)
            parameters["host"] = JsonSerializer.SerializeToElement(sourceHost);
        if (request.StringParam("sourceInstanceId") is { Length: > 0 } sourceInstanceId)
            parameters["instanceId"] = JsonSerializer.SerializeToElement(sourceInstanceId);
        return request with
        {
            App = app,
            FilePath = filePath,
            Target = null,
            Params = JsonSerializer.SerializeToElement(parameters),
        };
    }
}
