using System.Globalization;
using System.Text.Json;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Runtime;

internal static class JsonParameters
{
    public static string RequiredString(this JsonElement parameters, string name)
    {
        var value = parameters.OptionalString(name);
        return string.IsNullOrWhiteSpace(value)
            ? throw new OfficeWorkerException("invalid_params", $"缺少参数: {name}")
            : value;
    }

    public static string? OptionalString(this JsonElement parameters, string name)
    {
        if (parameters.ValueKind != JsonValueKind.Object || !parameters.TryGetProperty(name, out var value))
        {
            return null;
        }

        return value.ValueKind == JsonValueKind.String ? value.GetString() : null;
    }

    public static bool OptionalBoolean(this JsonElement parameters, string name, bool defaultValue = false)
    {
        if (parameters.ValueKind != JsonValueKind.Object || !parameters.TryGetProperty(name, out var value))
        {
            return defaultValue;
        }

        return value.ValueKind is JsonValueKind.True or JsonValueKind.False ? value.GetBoolean() : defaultValue;
    }

    public static int OptionalInt32(this JsonElement parameters, string name, int defaultValue = 0)
    {
        if (parameters.ValueKind != JsonValueKind.Object || !parameters.TryGetProperty(name, out var value))
        {
            return defaultValue;
        }

        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var number))
        {
            return number;
        }

        return int.TryParse(value.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out number)
            ? number
            : defaultValue;
    }

    public static double OptionalDouble(this JsonElement parameters, string name, double defaultValue = 0)
    {
        if (parameters.ValueKind != JsonValueKind.Object || !parameters.TryGetProperty(name, out var value))
        {
            return defaultValue;
        }

        return value.ValueKind == JsonValueKind.Number && value.TryGetDouble(out var number)
            ? number
            : defaultValue;
    }

    public static JsonElement PropertyOrEmpty(this JsonElement parameters, string name)
    {
        return parameters.ValueKind == JsonValueKind.Object && parameters.TryGetProperty(name, out var value)
            ? value
            : default;
    }
}
