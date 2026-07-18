using System.Text.Json;
using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Excel;

internal sealed class ExcelRangeService(ExcelSessionService sessions)
{
    private sealed record RangeSnapshot(object? Content);

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

            var singleFormula = matrix.GetLength(0) == 1 && matrix.GetLength(1) == 1
                ? Convert.ToString(matrix[0, 0])?.Trim()
                : null;
            var isMultiCellLegacyCse = legacyCse
                && !string.IsNullOrWhiteSpace(singleFormula)
                && singleFormula[0] == '='
                && Convert.ToInt32(startApi.Rows.Count) * Convert.ToInt32(startApi.Columns.Count) > 1;
            if (isMultiCellLegacyCse)
            {
                ExcelRangeWriteTransaction.Execute(
                    () => CaptureSnapshot(startApi),
                    () => ExcelFormulaWriter.Write(
                        new ComExcelFormulaCell(startRange, readFormulaArray: true),
                        singleFormula!,
                        legacyCse: true),
                    snapshot => RestoreSnapshot(startApi, snapshot));

                return new
                {
                    written = matrix.Length,
                    dynamicCells = 0,
                    arrayCells = 1,
                    plainCells = 0,
                };
            }

            // WPS 的旧版当前窗口链路通过单元格默认 Value 入口写入，
            // 由 WPS 自己解析以 '=' 开头的文本。不要把 WPS 的 range.write
            // 重新送进 Formula/Formula2，否则动态数组会走另一套语义。
            if (string.Equals(handle.ProgId, "Ket.Application", StringComparison.OrdinalIgnoreCase))
            {
                targetRange = startApi.Resize[matrix.GetLength(0), matrix.GetLength(1)];
                dynamic wpsTargetApi = targetRange;
                var wpsPlan = ExcelRangeWritePlan.Create(matrix, legacyCse);
                ExcelRangeWriteTransaction.Execute(
                    () => CaptureSnapshot(wpsTargetApi),
                    () =>
                    {
                        for (var row = 0; row < matrix.GetLength(0); row++)
                        {
                            for (var column = 0; column < matrix.GetLength(1); column++)
                            {
                                object? cell = null;
                                try
                                {
                                    cell = wpsTargetApi.Cells.Item(row + 1, column + 1);
                                    dynamic cellApi = cell;
                                    cellApi.Value = matrix[row, column];
                                }
                                finally
                                {
                                    ComInterop.Release(cell);
                                }
                            }
                        }
                    },
                    snapshot => RestoreSnapshot(wpsTargetApi, snapshot));

                return new
                {
                    written = matrix.Length,
                    dynamicCells = wpsPlan.DynamicCells,
                    arrayCells = wpsPlan.ArrayCells,
                    plainCells = wpsPlan.PlainCells,
                };
            }

            targetRange = startApi.Resize[matrix.GetLength(0), matrix.GetLength(1)];
            dynamic targetApi = targetRange;

            var plan = ExcelRangeWritePlan.Create(matrix, legacyCse);
            ExcelRangeWriteTransaction.Execute(
                () => CaptureSnapshot(targetApi),
                () =>
                {
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
                },
                snapshot => RestoreSnapshot(targetApi, snapshot));

            return new
            {
                written = matrix.Length,
                dynamicCells = plan.DynamicCells,
                arrayCells = plan.ArrayCells,
                plainCells = plan.PlainCells,
            };
        }
        catch (OfficeWorkerException exception) when (exception.Code is "formula_rejected" or "legacy_array_unsupported")
        {
            var host = handle.ProgId == "Ket.Application" ? "WPS 表格" : "Microsoft Excel";
            var version = "unknown";
            var build = "unknown";
            try { version = Convert.ToString(app.Version) ?? "unknown"; } catch { }
            try { build = Convert.ToString(app.Build) ?? "unknown"; } catch { }
            var diagnosis = exception.Code == "formula_rejected"
                ? "当前宿主拒绝了 Formula 写入；请检查函数版本、区域语法和公式本地化"
                : "当前宿主不支持传统 FormulaArray 写入";
            throw new OfficeWorkerException(
                exception.Code,
                $"{exception.Message}。诊断：{diagnosis}。当前绑定宿主：{host}（{handle.ProgId}，版本 {version}，构建 {build}）。如果目标是另一个程序，请先在连接状态中选择对应宿主。",
                exception.Details,
                exception);
        }
        finally
        {
            ComInterop.Release(targetRange);
            ComInterop.Release(startRange);
            ComInterop.Release(sheet);
            ComInterop.Release(workbook);
        }
    }

    private static RangeSnapshot CaptureSnapshot(dynamic target)
    {
        try
        {
            return new RangeSnapshot(target.Formula);
        }
        catch
        {
            return new RangeSnapshot(target.Formula);
        }
    }

    private static void RestoreSnapshot(dynamic target, RangeSnapshot snapshot)
    {
        target.Formula = snapshot.Content;
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
