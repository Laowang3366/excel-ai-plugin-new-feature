using System.Text.Json;
using Wengge.OfficeWorker.Com;

namespace Wengge.OfficeWorker.Excel;

internal sealed class ExcelRangeService(ExcelSessionService sessions)
{
    public object Read(string sheetName, string address, string expand)
    {
        using var handle = sessions.GetActiveRequired();
        dynamic app = handle.Application;
        object? workbook = null;
        object? sheet = null;
        object? sourceRange = null;
        object? readRange = null;
        try
        {
            workbook = app.ActiveWorkbook ?? throw new InvalidOperationException("当前没有活动工作簿");
            dynamic workBookApi = workbook;
            sheet = workBookApi.Sheets.Item(sheetName);
            dynamic sheetApi = sheet;
            sourceRange = sheetApi.Range(address);
            readRange = ExpandRange(sourceRange, expand, out var expanded);
            dynamic rangeApi = readRange;
            return new
            {
                values = ExcelValueConverter.ToRows(rangeApi.Value2),
                address = Convert.ToString(rangeApi.Address[false, false, 1, false]),
                expanded,
                expandMode = expand,
            };
        }
        finally
        {
            if (!ReferenceEquals(readRange, sourceRange)) ComInterop.Release(readRange);
            ComInterop.Release(sourceRange);
            ComInterop.Release(sheet);
            ComInterop.Release(workbook);
        }
    }

    public object Write(string sheetName, string address, JsonElement values, bool legacyCse = false)
    {
        var matrix = ExcelValueConverter.FromJsonRows(values);
        if (matrix.Length == 0)
        {
            return new { written = 0, dynamicCells = 0, arrayCells = 0, plainCells = 0 };
        }

        using var handle = sessions.GetActiveRequired();
        dynamic app = handle.Application;
        object? workbook = null;
        object? sheet = null;
        object? startRange = null;
        object? targetRange = null;
        try
        {
            workbook = app.ActiveWorkbook ?? throw new InvalidOperationException("当前没有活动工作簿");
            dynamic workBookApi = workbook;
            sheet = workBookApi.Sheets.Item(sheetName);
            dynamic sheetApi = sheet;
            startRange = sheetApi.Range(address);
            dynamic startApi = startRange;
            targetRange = startApi.Resize[matrix.GetLength(0), matrix.GetLength(1)];
            dynamic targetApi = targetRange;

            var plan = ExcelRangeWritePlan.Create(matrix, legacyCse);
            targetApi.Value2 = plan.BulkValues;
            foreach (var formula in plan.Formulas)
            {
                object? cell = null;
                try
                {
                    cell = targetApi.Cells.Item(formula.Row + 1, formula.Column + 1);
                    ExcelFormulaWriter.Write(new ComExcelFormulaCell(cell), formula.Formula, legacyCse);
                }
                finally
                {
                    ComInterop.Release(cell);
                }
            }

            return new
            {
                written = matrix.Length,
                dynamicCells = plan.DynamicCells,
                arrayCells = plan.ArrayCells,
                plainCells = plan.PlainCells,
            };
        }
        finally
        {
            ComInterop.Release(targetRange);
            ComInterop.Release(startRange);
            ComInterop.Release(sheet);
            ComInterop.Release(workbook);
        }
    }

    public object Clear(string sheetName, string address)
    {
        using var handle = sessions.GetActiveRequired();
        dynamic app = handle.Application;
        object? workbook = null;
        object? sheet = null;
        object? range = null;
        try
        {
            workbook = app.ActiveWorkbook ?? throw new InvalidOperationException("当前没有活动工作簿");
            dynamic workBookApi = workbook;
            sheet = workBookApi.Sheets.Item(sheetName);
            dynamic sheetApi = sheet;
            range = sheetApi.Range(address);
            dynamic rangeApi = range;
            rangeApi.Clear();
            return new { cleared = true };
        }
        finally
        {
            ComInterop.Release(range);
            ComInterop.Release(sheet);
            ComInterop.Release(workbook);
        }
    }

    public object GetSelection(bool includeValues)
    {
        using var handle = sessions.GetActiveRequired();
        dynamic app = handle.Application;
        object? selection = null;
        object? worksheet = null;
        try
        {
            selection = app.Selection ?? throw new InvalidOperationException("当前没有可用选区");
            dynamic selectionApi = selection;
            worksheet = selectionApi.Worksheet;
            dynamic worksheetApi = worksheet;
            var address = Convert.ToString(selectionApi.Address[false, false, 1, false]);
            var sheetName = Convert.ToString(worksheetApi.Name);
            return includeValues
                ? new { address, sheetName, values = ExcelValueConverter.ToRows(selectionApi.Value2) }
                : new { address, sheetName };
        }
        finally
        {
            ComInterop.Release(worksheet);
            ComInterop.Release(selection);
        }
    }

    private static object ExpandRange(object source, string mode, out bool expanded)
    {
        expanded = false;
        dynamic range = source;
        object? candidate = null;
        try
        {
            candidate = mode.ToLowerInvariant() switch
            {
                "spill" => TrySpillRange(range),
                "currentarray" => range.CurrentArray,
                "currentregion" => range.CurrentRegion,
                _ => null,
            };
            if (candidate is null)
            {
                return source;
            }

            dynamic candidateApi = candidate;
            expanded = Convert.ToString(candidateApi.Address[false, false, 1, false]) !=
                Convert.ToString(range.Address[false, false, 1, false]);
            return candidate;
        }
        catch
        {
            ComInterop.Release(candidate);
            return source;
        }
    }

    private static object? TrySpillRange(dynamic range)
    {
        object? firstCell = null;
        try
        {
            firstCell = range.Cells.Item(1, 1);
            dynamic cellApi = firstCell;
            try
            {
                return cellApi.SpillingToRange;
            }
            catch
            {
                return cellApi.CurrentArray;
            }
        }
        finally
        {
            ComInterop.Release(firstCell);
        }
    }
}
