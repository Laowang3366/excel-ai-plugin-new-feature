using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Office;

internal sealed class ExcelActionService(
    OfficeApplicationProvider applications,
    ExcelQueryActionService queries,
    ExcelWorkbookActionService workbookActions,
    ExcelFormulaActionService formulaActions,
    ExcelCrossOfficeActionService crossOffice)
{
    public object Execute(OfficeActionRequest request)
    {
        if (ExcelQueryActionService.Supports(request.Operation)) return queries.Execute(request);
        if (ExcelWorkbookActionService.Supports(request.Operation)) return workbookActions.Execute(request);
        if (ExcelFormulaActionService.Supports(request.Operation)) return formulaActions.Execute(request);
        if (ExcelCrossOfficeActionService.Supports(request.Operation)) return crossOffice.Execute(request);
        using var context = new ExcelActionContext(applications, request);
        var result = request.Operation switch
        {
            "insertChart" => InsertChart(context, request),
            "setDataValidation" => SetDataValidation(context, request),
            "applyConditionalFormatting" => ApplyConditionalFormatting(context, request),
            "styleTable" => StyleTable(context, request),
            "createPivotTable" => CreatePivotTable(context, request),
            "refreshPivotTables" => RefreshPivotTables(context, request),
            "addSlicer" => AddSlicer(context, request),
            "exportPdf" or "exportSheetsToPdf" => ExportPdf(context, request),
            "snapshot" => Snapshot(context, request),
            _ => throw new OfficeWorkerException("unsupported_operation", $"不支持的 Excel COM 操作: {request.Operation}"),
        };
        if (request.Operation is not ("exportPdf" or "exportSheetsToPdf" or "snapshot")) context.Save(request);
        return result;
    }

    private static object InsertChart(ExcelActionContext context, OfficeActionRequest request)
    {
        var (sheet, range) = context.GetRange(request);
        object? chartObjects = null;
        object? chartObject = null;
        object? chart = null;
        try
        {
            dynamic sheetApi = sheet;
            dynamic rangeApi = range;
            chartObjects = sheetApi.ChartObjects();
            dynamic chartObjectsApi = chartObjects;
            chartObject = chartObjectsApi.Add(
                Convert.ToDouble(rangeApi.Left) + Convert.ToDouble(rangeApi.Width) + 20,
                Convert.ToDouble(rangeApi.Top), 420, 260);
            chart = ((dynamic)chartObject).Chart;
            dynamic chartApi = chart;
            chartApi.SetSourceData(range);
            chartApi.ChartType = ChartType(request.StringParam("chartType"));
            return Done(request, "chart", "已插入 Excel 图表", new { chartType = request.StringParam("chartType", "column") });
        }
        finally
        {
            ComInterop.Release(chart);
            ComInterop.Release(chartObject);
            ComInterop.Release(chartObjects);
            ComInterop.Release(range);
            ComInterop.Release(sheet);
        }
    }

    private static object SetDataValidation(ExcelActionContext context, OfficeActionRequest request)
    {
        var formula = request.StringParam("formula", request.StringParam("list"));
        if (string.IsNullOrWhiteSpace(formula)) throw new OfficeWorkerException("invalid_params", "setDataValidation 需要 params.formula 或 params.list");
        var (sheet, range) = context.GetRange(request);
        object? validation = null;
        try
        {
            validation = ((dynamic)range).Validation;
            dynamic validationApi = validation;
            try { validationApi.Delete(); } catch { }
            validationApi.Add(3, 1, 1, formula);
            validationApi.IgnoreBlank = true;
            validationApi.InCellDropdown = true;
            return Done(request, "validation", "已设置数据验证", new { formula });
        }
        finally
        {
            ComInterop.Release(validation);
            ComInterop.Release(range);
            ComInterop.Release(sheet);
        }
    }

    private static object ApplyConditionalFormatting(ExcelActionContext context, OfficeActionRequest request)
    {
        var formula = request.StringParam("formula", "TRUE");
        var (sheet, range) = context.GetRange(request);
        object? conditions = null;
        object? condition = null;
        object? interior = null;
        try
        {
            conditions = ((dynamic)range).FormatConditions;
            condition = ((dynamic)conditions).Add(2, 3, formula);
            interior = ((dynamic)condition).Interior;
            ((dynamic)interior).Color = OleColor(request.StringParam("fillColor", "FFF2CC"));
            return Done(request, "conditional-format", "已设置条件格式", new { formula });
        }
        finally
        {
            ComInterop.Release(interior);
            ComInterop.Release(condition);
            ComInterop.Release(conditions);
            ComInterop.Release(range);
            ComInterop.Release(sheet);
        }
    }

    private static object StyleTable(ExcelActionContext context, OfficeActionRequest request)
    {
        var (sheet, range) = context.GetRange(request);
        object? font = null;
        object? borders = null;
        object? interior = null;
        try
        {
            dynamic rangeApi = range;
            font = rangeApi.Font;
            ((dynamic)font).Bold = true;
            borders = rangeApi.Borders;
            ((dynamic)borders).LineStyle = 1;
            interior = rangeApi.Interior;
            ((dynamic)interior).Color = OleColor(request.StringParam("headerColor", "1F4E79"));
            return Done(request, "style", "已应用 Excel 表格样式", null);
        }
        finally
        {
            ComInterop.Release(interior);
            ComInterop.Release(borders);
            ComInterop.Release(font);
            ComInterop.Release(range);
            ComInterop.Release(sheet);
        }
    }

    private static object CreatePivotTable(ExcelActionContext context, OfficeActionRequest request)
    {
        var (sheet, range) = context.GetRange(request);
        object? caches = null;
        object? cache = null;
        object? destinationSheet = null;
        object? destinationRange = null;
        object? pivot = null;
        try
        {
            var name = request.StringParam("name", $"AI_Pivot_{DateTime.Now:HHmmss}");
            var destination = request.StringParam("destination", $"{request.ExcelTarget().SheetName}!H3");
            var parts = destination.Split('!', 2);
            destinationSheet = parts.Length == 2 ? context.Workbook.Worksheets.Item(parts[0].Trim('\'')) : sheet;
            dynamic destinationSheetApi = destinationSheet;
            destinationRange = destinationSheetApi.Range(parts.Length == 2 ? parts[1] : destination);
            dynamic rangeApi = range;
            var sourceAddress = rangeApi.Address[true, true, 1, true];
            caches = context.Workbook.PivotCaches();
            cache = ((dynamic)caches).Create(1, sourceAddress);
            pivot = ((dynamic)cache).CreatePivotTable(destinationRange, name);
            dynamic pivotApi = pivot;
            SetPivotOrientations(pivotApi, request.Param("rowFields"), 1);
            SetPivotOrientations(pivotApi, request.Param("columnFields"), 2);
            SetPivotOrientations(pivotApi, request.Param("filterFields"), 3);
            AddPivotDataFields(pivotApi, request.Param("dataFields"));
            return Done(request, "pivot-table", "已创建数据透视表", new { pivotName = name, destination });
        }
        finally
        {
            ComInterop.Release(pivot);
            ComInterop.Release(destinationRange);
            if (!ReferenceEquals(destinationSheet, sheet)) ComInterop.Release(destinationSheet);
            ComInterop.Release(cache);
            ComInterop.Release(caches);
            ComInterop.Release(range);
            ComInterop.Release(sheet);
        }
    }

    private static object RefreshPivotTables(ExcelActionContext context, OfficeActionRequest request)
    {
        var refreshed = 0;
        object? worksheets = null;
        try
        {
            worksheets = context.Workbook.Worksheets;
            dynamic worksheetsApi = worksheets;
            for (var sheetIndex = 1; sheetIndex <= Convert.ToInt32(worksheetsApi.Count); sheetIndex++)
            {
                object? sheet = null;
                object? pivots = null;
                try
                {
                    sheet = worksheetsApi.Item(sheetIndex);
                    pivots = ((dynamic)sheet).PivotTables();
                    dynamic pivotsApi = pivots;
                    for (var index = 1; index <= Convert.ToInt32(pivotsApi.Count); index++)
                    {
                        object? pivot = null;
                        try { pivot = pivotsApi.Item(index); ((dynamic)pivot).RefreshTable(); refreshed++; }
                        finally { ComInterop.Release(pivot); }
                    }
                }
                finally { ComInterop.Release(pivots); ComInterop.Release(sheet); }
            }
            if (request.BoolParam("refreshConnections", true)) context.Workbook.RefreshAll();
            return Done(request, "pivot-refresh", "已刷新数据透视表", new { refreshed });
        }
        finally { ComInterop.Release(worksheets); }
    }

    private static object AddSlicer(ExcelActionContext context, OfficeActionRequest request)
    {
        var field = request.StringParam("field");
        if (field.Length == 0) throw new OfficeWorkerException("invalid_params", "addSlicer 需要 params.field");
        var (sheet, range) = context.GetRange(request);
        object? pivot = FindPivot(context.Workbook, request.StringParam("pivotName"));
        object? slicerCaches = null;
        object? cache = null;
        object? slicers = null;
        object? slicer = null;
        try
        {
            if (pivot is null) throw new OfficeWorkerException("pivot_not_found", "找不到可用于切片器的数据透视表");
            slicerCaches = context.Workbook.SlicerCaches;
            dynamic slicerCachesApi = slicerCaches;
            try { cache = slicerCachesApi.Add2(pivot, field); } catch { cache = slicerCachesApi.Add(pivot, field); }
            slicers = ((dynamic)cache).Slicers;
            dynamic rangeApi = range;
            var name = request.StringParam("name", $"AI_Slicer_{DateTime.Now:HHmmss}");
            slicer = ((dynamic)slicers).Add(sheet, Type.Missing, name, request.StringParam("caption", field),
                request.DoubleParam("top", Convert.ToDouble(rangeApi.Top)),
                request.DoubleParam("left", Convert.ToDouble(rangeApi.Left) + Convert.ToDouble(rangeApi.Width) + 20), 144, 180);
            return Done(request, "slicer", "已创建切片器", new { slicerName = name });
        }
        finally
        {
            ComInterop.Release(slicer);
            ComInterop.Release(slicers);
            ComInterop.Release(cache);
            ComInterop.Release(slicerCaches);
            ComInterop.Release(pivot);
            ComInterop.Release(range);
            ComInterop.Release(sheet);
        }
    }

    private static object ExportPdf(ExcelActionContext context, OfficeActionRequest request)
    {
        var output = OutputPath(request, "export.pdf");
        if (request.Operation == "exportSheetsToPdf")
            return ExportSheetsPdf(context, request, output);
        if (request.StringParam("scope") == "sheet")
        {
            var (sheet, range) = context.GetRange(request);
            try { ((dynamic)sheet).ExportAsFixedFormat(0, output); }
            finally { ComInterop.Release(range); ComInterop.Release(sheet); }
        }
        else context.Workbook.ExportAsFixedFormat(0, output);
        return Done(request, "export", "已导出 Excel PDF", new { outputPath = output }, output);
    }

    private static object ExportSheetsPdf(ExcelActionContext context, OfficeActionRequest request, string output)
    {
        var requestedNames = request.Param("sheetNames").ValueKind == System.Text.Json.JsonValueKind.Array
            ? request.Param("sheetNames").EnumerateArray()
                .Where(value => value.ValueKind == System.Text.Json.JsonValueKind.String && !string.IsNullOrWhiteSpace(value.GetString()))
                .Select(value => value.GetString()!)
                .ToArray()
            : [];
        var targets = new List<object>();
        var sheetNames = new List<string>();
        object? worksheets = null;
        try
        {
            worksheets = context.Workbook.Worksheets;
            dynamic worksheetsApi = worksheets;
            if (requestedNames.Length > 0)
            {
                foreach (var name in requestedNames)
                {
                    object? sheet = null;
                    try { sheet = worksheetsApi.Item(name); }
                    catch { throw new OfficeWorkerException("sheet_not_found", $"找不到导出工作表: {name}"); }
                    targets.Add(sheet);
                    sheetNames.Add(Convert.ToString(((dynamic)sheet).Name) ?? name);
                }
            }
            else
            {
                for (var index = 1; index <= Convert.ToInt32(worksheetsApi.Count); index++)
                {
                    var sheet = worksheetsApi.Item(index);
                    targets.Add(sheet);
                    sheetNames.Add(Convert.ToString(((dynamic)sheet).Name) ?? $"Sheet{index}");
                }
            }

            var mode = request.StringParam("mode", "combined");
            var outputs = new List<string>();
            if (mode == "separate")
            {
                var directory = request.StringParam("outputDirectory");
                if (directory.Length == 0) directory = Path.GetDirectoryName(output) ?? Environment.CurrentDirectory;
                directory = Path.GetFullPath(directory);
                Directory.CreateDirectory(directory);
                for (var index = 0; index < targets.Count; index++)
                {
                    var file = Path.Combine(directory, SafeFileName(sheetNames[index]) + ".pdf");
                    EnsureCanOverwrite(file, request);
                    ((dynamic)targets[index]).ExportAsFixedFormat(0, file);
                    outputs.Add(file);
                }
                output = directory;
            }
            else
            {
                EnsureCanOverwrite(output, request);
                for (var index = 0; index < targets.Count; index++) ((dynamic)targets[index]).Select(index == 0);
                object? activeSheet = null;
                try
                {
                    activeSheet = context.App.ActiveSheet;
                    ((dynamic)activeSheet).ExportAsFixedFormat(0, output);
                }
                finally { ComInterop.Release(activeSheet); }
                outputs.Add(output);
                mode = "combined";
            }
            return Done(request, "export", "已批量导出 Excel 工作表 PDF", new { outputPaths = outputs, mode, sheetNames }, output);
        }
        finally
        {
            foreach (var target in targets) ComInterop.Release(target);
            ComInterop.Release(worksheets);
        }
    }

    private static void EnsureCanOverwrite(string path, OfficeActionRequest request)
    {
        if (File.Exists(path) && !request.BoolParam("overwrite"))
            throw new OfficeWorkerException("file_exists", $"PDF 已存在，请设置 overwrite=true: {path}");
    }

    private static string SafeFileName(string value)
    {
        var invalid = Path.GetInvalidFileNameChars().ToHashSet();
        var result = new string(value.Select(character => invalid.Contains(character) ? '_' : character).ToArray());
        return string.IsNullOrWhiteSpace(result) ? "sheet" : result;
    }

    private static object Snapshot(ExcelActionContext context, OfficeActionRequest request)
    {
        var output = OutputPath(request, "snapshot.png");
        var (sheet, range) = context.GetRange(request);
        object? chartObjects = null;
        object? chartObject = null;
        object? chart = null;
        try
        {
            dynamic rangeApi = range;
            rangeApi.CopyPicture(1, 2);
            chartObjects = ((dynamic)sheet).ChartObjects();
            chartObject = ((dynamic)chartObjects).Add(rangeApi.Left, rangeApi.Top, Math.Max(1d, Convert.ToDouble(rangeApi.Width)), Math.Max(1d, Convert.ToDouble(rangeApi.Height)));
            chart = ((dynamic)chartObject).Chart;
            ((dynamic)chart).Paste();
            ((dynamic)chart).Export(output, "PNG");
            return Done(request, "snapshot", "已导出 Excel 区域快照", new { outputPath = output }, output);
        }
        finally
        {
            try { ((dynamic?)chartObject)?.Delete(); } catch { }
            ComInterop.Release(chart);
            ComInterop.Release(chartObject);
            ComInterop.Release(chartObjects);
            ComInterop.Release(range);
            ComInterop.Release(sheet);
        }
    }

    private static object Done(OfficeActionRequest request, string kind, string summary, object? data, string? outputPath = null) =>
        OfficeActionResults.Done(request, "com", summary, data, [new OfficeChange(kind, request.Target ?? request.Operation, summary)], outputPath);

    private static string OutputPath(OfficeActionRequest request, string fallbackName)
    {
        var output = request.OutputPath;
        if (string.IsNullOrWhiteSpace(output))
        {
            var source = request.FilePath ?? Environment.CurrentDirectory;
            output = Path.Combine(Path.GetDirectoryName(source) ?? Environment.CurrentDirectory, fallbackName);
        }
        output = Path.GetFullPath(output);
        Directory.CreateDirectory(Path.GetDirectoryName(output) ?? Environment.CurrentDirectory);
        return output;
    }

    private static int ChartType(string type) => type.ToLowerInvariant() switch
    {
        "line" => 4,
        "pie" => 5,
        "bar" => 57,
        "area" => 1,
        "scatter" => -4169,
        _ => 51,
    };

    internal static int OleColor(string hex)
    {
        var value = hex.Trim().TrimStart('#');
        if (value.Length != 6 || !int.TryParse(value, System.Globalization.NumberStyles.HexNumber, null, out var rgb)) return 0;
        var red = (rgb >> 16) & 255;
        var green = (rgb >> 8) & 255;
        var blue = rgb & 255;
        return red | (green << 8) | (blue << 16);
    }

    private static void SetPivotOrientations(dynamic pivot, System.Text.Json.JsonElement fields, int orientation)
    {
        if (fields.ValueKind != System.Text.Json.JsonValueKind.Array) return;
        foreach (var field in fields.EnumerateArray())
        {
            var name = field.GetString();
            if (string.IsNullOrWhiteSpace(name)) continue;
            object? pivotField = null;
            try { pivotField = pivot.PivotFields(name); ((dynamic)pivotField).Orientation = orientation; }
            finally { ComInterop.Release(pivotField); }
        }
    }

    private static void AddPivotDataFields(dynamic pivot, System.Text.Json.JsonElement fields)
    {
        if (fields.ValueKind != System.Text.Json.JsonValueKind.Array) return;
        foreach (var field in fields.EnumerateArray())
        {
            var name = field.ValueKind == System.Text.Json.JsonValueKind.String ? field.GetString() : field.TryGetProperty("name", out var nameValue) ? nameValue.GetString() : null;
            if (string.IsNullOrWhiteSpace(name)) continue;
            var caption = field.ValueKind == System.Text.Json.JsonValueKind.Object && field.TryGetProperty("caption", out var captionValue) ? captionValue.GetString() : $"汇总项: {name}";
            var function = field.ValueKind == System.Text.Json.JsonValueKind.Object && field.TryGetProperty("function", out var functionValue) ? functionValue.GetString() : "sum";
            var aggregate = function switch { "average" => -4106, "count" => -4112, "max" => -4136, "min" => -4139, _ => -4157 };
            object? pivotField = null;
            try { pivotField = pivot.PivotFields(name); pivot.AddDataField(pivotField, caption, aggregate); }
            finally { ComInterop.Release(pivotField); }
        }
    }

    private static object? FindPivot(dynamic workbook, string name)
    {
        object? worksheets = null;
        try
        {
            worksheets = workbook.Worksheets;
            dynamic worksheetsApi = worksheets;
            for (var sheetIndex = 1; sheetIndex <= Convert.ToInt32(worksheetsApi.Count); sheetIndex++)
            {
                object? sheet = null;
                object? pivots = null;
                try
                {
                    sheet = worksheetsApi.Item(sheetIndex);
                    pivots = ((dynamic)sheet).PivotTables();
                    dynamic pivotsApi = pivots;
                    for (var index = 1; index <= Convert.ToInt32(pivotsApi.Count); index++)
                    {
                        object? pivot = pivotsApi.Item(index);
                        if (string.IsNullOrWhiteSpace(name) || string.Equals(Convert.ToString(((dynamic)pivot).Name), name, StringComparison.OrdinalIgnoreCase)) return pivot;
                        ComInterop.Release(pivot);
                    }
                }
                finally { ComInterop.Release(pivots); ComInterop.Release(sheet); }
            }
            return null;
        }
        finally { ComInterop.Release(worksheets); }
    }
}
