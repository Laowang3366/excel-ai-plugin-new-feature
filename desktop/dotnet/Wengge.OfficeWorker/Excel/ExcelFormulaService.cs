using Wengge.OfficeWorker.Com;

namespace Wengge.OfficeWorker.Excel;

internal sealed class ExcelFormulaService(ExcelSessionService sessions)
{
    public object GetContext(string sheetName, string? address)
    {
        using var handle = sessions.GetActiveRequired();
        dynamic app = handle.Application;
        object? workbook = null;
        object? sheet = null;
        object? range = null;
        try
        {
            workbook = app.ActiveWorkbook ?? throw new InvalidOperationException("当前没有活动工作簿");
            dynamic workbookApi = workbook;
            sheet = workbookApi.Sheets.Item(sheetName);
            dynamic sheetApi = sheet;
            range = string.IsNullOrWhiteSpace(address) ? sheetApi.UsedRange : sheetApi.Range(address);
            dynamic rangeApi = range;

            // Quarantine COM dynamic payloads at the acquisition boundary so ToRows/indexing/LINQ stay static.
            object? formulaMatrix = TryGetFormula(rangeApi);
            object? valueMatrix = rangeApi.Value2;
            object?[][] formulaRows = ExcelValueConverter.ToRows(formulaMatrix);
            object?[][] values = ExcelValueConverter.ToRows(valueMatrix);
            string rangeAddress =
                Convert.ToString(rangeApi.Address[false, false, 1, false]) ?? address ?? sheetName;

            var items = new List<object>();
            for (var row = 0; row < formulaRows.Length; row++)
            {
                for (var column = 0; column < formulaRows[row].Length; column++)
                {
                    // Force static object/string before prefix checks: COM Formula2 rows may stay dynamic
                    // and `dynamic.StartsWith('=', StringComparison)` throws RuntimeBinderException.
                    if (!TryGetFormulaText(formulaRows[row][column], out string formula))
                    {
                        continue;
                    }

                    object? cell = null;
                    try
                    {
                        cell = rangeApi.Cells.Item(row + 1, column + 1);
                        dynamic cellApi = cell;
                        string cellAddress =
                            Convert.ToString(cellApi.Address[false, false, 1, false]) ?? string.Empty;
                        object? cellValue =
                            row < values.Length && column < values[row].Length
                                ? values[row][column]
                                : null;
                        items.Add(
                            new
                            {
                                address = cellAddress,
                                formula,
                                value = cellValue,
                            });
                    }
                    finally
                    {
                        ComInterop.Release(cell);
                    }
                }
            }

            return new
            {
                sheetName,
                address = rangeAddress,
                formulas = items,
            };
        }
        finally
        {
            ComInterop.Release(range);
            ComInterop.Release(sheet);
            ComInterop.Release(workbook);
        }
    }

    private static object? TryGetFormula(dynamic range)
    {
        try
        {
            // Return as object so callers never keep a dynamic Formula2/Formula payload.
            object? formula2 = range.Formula2;
            return formula2;
        }
        catch
        {
            object? formula = range.Formula;
            return formula;
        }
    }

    /// <summary>
    /// Convert a COM/dynamic cell value to a static formula string and detect leading '='.
    /// Must not call StartsWith on a dynamic receiver with (char, StringComparison).
    /// </summary>
    internal static bool TryGetFormulaText(object? cellValue, out string formula)
    {
        // Cast through object so Convert.ToString returns a real string, not a dynamic.
        formula = Convert.ToString(cellValue is null ? null : (object)cellValue) ?? string.Empty;
        // Use string.StartsWith("=") on the static string — never char overload on dynamic.
        if (
            string.IsNullOrWhiteSpace(formula)
            || formula.Length < 2
            || !formula.StartsWith("=", StringComparison.Ordinal)
        )
        {
            formula = string.Empty;
            return false;
        }

        return true;
    }
}
