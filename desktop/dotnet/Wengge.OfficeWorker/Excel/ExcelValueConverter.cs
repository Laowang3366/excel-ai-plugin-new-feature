using System.Text.Json;

namespace Wengge.OfficeWorker.Excel;

internal static class ExcelValueConverter
{
    public static object?[][] ToRows(object? value)
    {
        if (value is not Array array)
        {
            return [[Normalize(value)]];
        }

        if (array.Rank == 1)
        {
            return Enumerable.Range(array.GetLowerBound(0), array.GetLength(0))
                .Select(index => new[] { Normalize(array.GetValue(index)) })
                .ToArray();
        }

        var rows = new object?[array.GetLength(0)][];
        for (var row = 0; row < rows.Length; row++)
        {
            rows[row] = new object?[array.GetLength(1)];
            for (var column = 0; column < rows[row].Length; column++)
            {
                rows[row][column] = Normalize(array.GetValue(row + array.GetLowerBound(0), column + array.GetLowerBound(1)));
            }
        }

        return rows;
    }

    public static object[,] FromJsonRows(JsonElement values)
    {
        if (values.ValueKind != JsonValueKind.Array)
        {
            throw new ArgumentException("values 必须是二维数组", nameof(values));
        }

        var rows = values.EnumerateArray().ToArray();
        if (rows.Length == 0)
        {
            return new object[0, 0];
        }

        var width = rows.Max(row => row.ValueKind == JsonValueKind.Array ? row.GetArrayLength() : 1);
        var result = new object[rows.Length, width];
        for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
        {
            var row = rows[rowIndex].ValueKind == JsonValueKind.Array
                ? rows[rowIndex].EnumerateArray().ToArray()
                : [rows[rowIndex]];
            for (var columnIndex = 0; columnIndex < width; columnIndex++)
            {
                result[rowIndex, columnIndex] = columnIndex < row.Length ? FromJsonValue(row[columnIndex]) ?? string.Empty : string.Empty;
            }
        }

        return result;
    }

    public static object? FromJsonValue(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.String => value.GetString(),
        JsonValueKind.Number when value.TryGetInt64(out var integer) => integer,
        JsonValueKind.Number => value.GetDouble(),
        JsonValueKind.True => true,
        JsonValueKind.False => false,
        JsonValueKind.Null or JsonValueKind.Undefined => null,
        _ => value.GetRawText(),
    };

    private static object Normalize(object? value) => value ?? string.Empty;
}
