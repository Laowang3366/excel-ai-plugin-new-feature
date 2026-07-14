using System.Globalization;
using System.Text.Json;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Office;

internal static class PresentationThemePackageUpdater
{
    public static bool Apply(string? filePath, JsonElement themeColors)
    {
        var updates = ReadUpdates(themeColors);
        if (updates.Count == 0) return false;
        if (string.IsNullOrWhiteSpace(filePath)) throw new OfficeWorkerException("invalid_params", "WPS 主题色兼容更新缺少演示文稿路径");
        var fullPath = System.IO.Path.GetFullPath(filePath);
        Exception? lastError = null;
        for (var attempt = 0; attempt < 10; attempt++)
        {
            try { return ApplyCore(fullPath, updates); }
            catch (IOException exception)
            {
                lastError = exception;
                Thread.Sleep(200);
            }
        }
        throw new OfficeWorkerException("file_locked", $"WPS 保存后演示文稿仍被占用: {fullPath}", null, lastError);
    }

    private static bool ApplyCore(string filePath, IReadOnlyDictionary<int, string> updates)
    {
        using var document = PresentationDocument.Open(filePath, true);
        var presentationPart = document.PresentationPart
            ?? throw new OfficeWorkerException("invalid_presentation", "演示文稿缺少 presentation 部件");
        var changed = false;
        var seen = new HashSet<Uri>();
        foreach (var masterPart in presentationPart.SlideMasterParts)
        {
            var themePart = masterPart.ThemePart;
            if (themePart?.Uri is null || !seen.Add(themePart.Uri)) continue;
            var theme = themePart.Theme;
            var scheme = theme?.ThemeElements?.ColorScheme;
            if (theme is null || scheme is null) continue;
            foreach (var (index, value) in updates)
            {
                if (index < 1 || index > scheme.ChildElements.Count) continue;
                var slot = scheme.ChildElements[index - 1];
                slot.RemoveAllChildren();
                slot.Append(new RgbColorModelHex { Val = value });
                changed = true;
            }
            theme.Save(themePart);
        }
        if (!changed) throw new OfficeWorkerException("theme_not_found", "演示文稿没有可更新的主题色方案");
        return true;
    }

    private static Dictionary<int, string> ReadUpdates(JsonElement themeColors)
    {
        var result = new Dictionary<int, string>();
        if (themeColors.ValueKind != JsonValueKind.Array) return result;
        foreach (var item in themeColors.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.Object || !item.TryGetProperty("index", out var indexValue) || !indexValue.TryGetInt32(out var index)
                || index < 1 || index > 12 || !item.TryGetProperty("value", out var colorValue)) continue;
            var color = colorValue.ValueKind == JsonValueKind.String ? colorValue.GetString()?.Trim().TrimStart('#') : null;
            if (color?.Length == 6 && uint.TryParse(color, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _))
                result[index] = color.ToUpperInvariant();
        }
        return result;
    }
}
