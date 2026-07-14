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
            var formulas = TryGetFormula(rangeApi);
            var values = ExcelValueConverter.ToRows(rangeApi.Value2);
            var formulaRows = ExcelValueConverter.ToRows(formulas);
            var items = new List<object>();
            for (var row = 0; row < formulaRows.Length; row++)
            {
                for (var column = 0; column < formulaRows[row].Length; column++)
                {
                    var formula = Convert.ToString(formulaRows[row][column]);
                    if (string.IsNullOrWhiteSpace(formula) || !formula.StartsWith('=', StringComparison.Ordinal))
                    {
                        continue;
                    }

                    object? cell = null;
                    try
                    {
                        cell = rangeApi.Cells.Item(row + 1, column + 1);
                        dynamic cellApi = cell;
                        items.Add(new
                        {
                            address = Convert.ToString(cellApi.Address[false, false, 1, false]),
                            formula,
                            value = values.ElementAtOrDefault(row)?.ElementAtOrDefault(column),
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
                address = Convert.ToString(rangeApi.Address[false, false, 1, false]),
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
            return range.Formula2;
        }
        catch
        {
            return range.Formula;
        }
    }
}
