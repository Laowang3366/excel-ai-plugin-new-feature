using System.Text.Json;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Office;

internal sealed record OfficeActionRequest(
    string App,
    string Action,
    string Operation,
    string? FilePath,
    string? OutputPath,
    string? Target,
    string? PreferEngine,
    JsonElement Params)
{
    public static OfficeActionRequest Parse(JsonElement value)
    {
        if (value.ValueKind != JsonValueKind.Object)
            throw new OfficeWorkerException("invalid_params", "Office action 必须是对象");
        return new OfficeActionRequest(
            Required(value, "app"),
            Required(value, "action"),
            Required(value, "operation"),
            Optional(value, "filePath"),
            Optional(value, "outputPath"),
            Optional(value, "target"),
            Optional(value, "preferEngine"),
            value.TryGetProperty("params", out var parameters) && parameters.ValueKind == JsonValueKind.Object
                ? parameters
                : JsonSerializer.SerializeToElement(new { }));
    }

    public string StringParam(string name, string fallback = "") =>
        Params.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String ? value.GetString() ?? fallback : fallback;

    public bool BoolParam(string name, bool fallback = false) =>
        Params.TryGetProperty(name, out var value) && value.ValueKind is JsonValueKind.True or JsonValueKind.False ? value.GetBoolean() : fallback;

    public int IntParam(string name, int fallback = 0) =>
        Params.TryGetProperty(name, out var value) && value.TryGetInt32(out var result) ? result : fallback;

    public double DoubleParam(string name, double fallback = 0) =>
        Params.TryGetProperty(name, out var value) && value.TryGetDouble(out var result) ? result : fallback;

    public JsonElement Param(string name) => Params.TryGetProperty(name, out var value) ? value : default;

    public (string SheetName, string Address) ExcelTarget()
    {
        var sheetName = StringParam("sheetName", "Sheet1");
        var address = StringParam("range", "A1");
        if (Target?.StartsWith("range:", StringComparison.OrdinalIgnoreCase) == true)
        {
            var locator = Target[6..];
            var separator = locator.IndexOf('!');
            if (separator >= 0)
            {
                sheetName = locator[..separator].Trim('\'');
                address = locator[(separator + 1)..];
            }
            else if (locator.Length > 0)
            {
                address = locator;
            }
        }
        return (sheetName, address);
    }

    public int SlideIndex()
    {
        if (Target?.StartsWith("slide:", StringComparison.OrdinalIgnoreCase) == true && int.TryParse(Target[6..], out var index))
            return Math.Max(1, index);
        return Math.Max(1, IntParam("slideIndex", 1));
    }

    private static string Required(JsonElement value, string name) =>
        Optional(value, name) is { Length: > 0 } text ? text : throw new OfficeWorkerException("invalid_params", $"缺少参数: {name}");

    private static string? Optional(JsonElement value, string name) =>
        value.TryGetProperty(name, out var property) && property.ValueKind == JsonValueKind.String ? property.GetString() : null;
}

internal sealed record OfficeChange(string Kind, string? Target, string Detail);

internal static class OfficeActionResults
{
    private static readonly JsonSerializerOptions MetadataSerializerOptions = new(JsonSerializerDefaults.Web)
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
    };

    public static object WithProgId(object data, string progId)
    {
        var value = JsonSerializer.SerializeToElement(data, MetadataSerializerOptions);
        if (value.ValueKind != JsonValueKind.Object) return new { progId, value = data };
        var result = value.EnumerateObject().ToDictionary(
            property => property.Name,
            property => (object?)property.Value.Clone(),
            StringComparer.Ordinal);
        result["progId"] = progId;
        return result;
    }

    public static object Done(OfficeActionRequest request, string engine, string summary, object? data, IEnumerable<OfficeChange>? changes = null, string? outputPath = null) => new
    {
        status = "done",
        engine,
        app = request.App,
        action = request.Action,
        operation = request.Operation,
        filePath = request.FilePath,
        outputPath = outputPath ?? request.OutputPath ?? request.FilePath,
        target = request.Target,
        summary,
        data,
        changes = changes?.Select(change => new { kind = change.Kind, target = change.Target, detail = change.Detail }).ToArray() ?? [],
    };

    public static object NeedsCom(OfficeActionRequest request, string summary) => new
    {
        status = "needsCom",
        engine = "openxml",
        app = request.App,
        action = request.Action,
        operation = request.Operation,
        filePath = request.FilePath,
        outputPath = request.OutputPath,
        target = request.Target,
        summary,
        changes = Array.Empty<object>(),
    };
}
