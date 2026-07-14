using System.Text.Json;
using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Office;

internal sealed class ExcelChartActionService(OfficeApplicationProvider applications)
{
    public static bool Supports(string operation) => operation is "inspectCharts" or "formatChart";

    public object Execute(OfficeActionRequest request)
    {
        using var context = new ExcelActionContext(applications, request);
        if (request.Operation == "inspectCharts")
            return OfficeActionResults.Done(request, "com", "已检查工作簿图表", InspectCharts(context.Workbook, request.StringParam("chartName")), Array.Empty<OfficeChange>());
        var snapshot = FormatChart(context, request);
        context.Save(request);
        return OfficeActionResults.Done(request, "com", "已完成图表深度编辑", snapshot,
            [new OfficeChange("chart-style", request.StringParam("chartName"), "已完成图表深度编辑")]);
    }

    private static object InspectCharts(dynamic workbook, string chartName)
    {
        var charts = new List<object>();
        object? worksheets = null;
        try
        {
            worksheets = workbook.Worksheets;
            dynamic worksheetsApi = worksheets;
            for (var sheetIndex = 1; sheetIndex <= Convert.ToInt32(worksheetsApi.Count); sheetIndex++)
            {
                object? sheet = null;
                object? chartObjects = null;
                try
                {
                    sheet = worksheetsApi.Item(sheetIndex);
                    dynamic sheetApi = sheet;
                    chartObjects = sheetApi.ChartObjects();
                    dynamic chartObjectsApi = chartObjects;
                    for (var index = 1; index <= Convert.ToInt32(chartObjectsApi.Count); index++)
                    {
                        object? chartObject = null;
                        try
                        {
                            chartObject = chartObjectsApi.Item(index);
                            dynamic chartObjectApi = chartObject;
                            var name = Convert.ToString(chartObjectApi.Name) ?? string.Empty;
                            if (chartName.Length > 0 && !name.Equals(chartName, StringComparison.OrdinalIgnoreCase)) continue;
                            charts.Add(Snapshot(chartObjectApi, Convert.ToString(sheetApi.Name) ?? string.Empty));
                        }
                        finally { ComInterop.Release(chartObject); }
                    }
                }
                finally { ComInterop.Release(chartObjects); ComInterop.Release(sheet); }
            }
            return new { charts, chartCount = charts.Count };
        }
        finally { ComInterop.Release(worksheets); }
    }

    private static object FormatChart(ExcelActionContext context, OfficeActionRequest request)
    {
        var (sheet, range) = context.GetRange(request);
        object? chartObjects = null;
        object? chartObject = null;
        object? chart = null;
        try
        {
            chartObjects = ((dynamic)sheet).ChartObjects();
            dynamic chartObjectsApi = chartObjects;
            var chartName = request.StringParam("chartName");
            chartObject = chartName.Length > 0
                ? chartObjectsApi.Item(chartName)
                : chartObjectsApi.Item(Math.Max(1, request.IntParam("chartIndex", 1)));
            dynamic chartObjectApi = chartObject;
            chart = chartObjectApi.Chart;
            dynamic chartApi = chart;
            var sourceRange = request.StringParam("sourceRange");
            if (sourceRange.Length > 0)
            {
                object? source = null;
                try { source = ResolveRange(context.Workbook, sheet, sourceRange); chartApi.SetSourceData(source); }
                finally { ComInterop.Release(source); }
            }
            var chartType = request.StringParam("chartType");
            if (chartType.Length > 0) chartApi.ChartType = ResolveChartType(chartType);
            if (request.Param("name").ValueKind == JsonValueKind.String) chartObjectApi.Name = request.StringParam("name");
            if (request.Param("showTitle").ValueKind is JsonValueKind.True or JsonValueKind.False) chartApi.HasTitle = request.BoolParam("showTitle");
            if (request.Param("title").ValueKind == JsonValueKind.String)
            {
                chartApi.HasTitle = true;
                object? title = null;
                try { title = chartApi.ChartTitle; ((dynamic)title).Text = request.StringParam("title"); }
                finally { ComInterop.Release(title); }
            }
            if (request.IntParam("style") > 0) { try { chartApi.ChartStyle = request.IntParam("style"); } catch { } }
            if (request.Param("showLegend").ValueKind is JsonValueKind.True or JsonValueKind.False) chartApi.HasLegend = request.BoolParam("showLegend");
            SetPosition(chartObjectApi, request);
            if (request.BoolParam("replaceSeries")) DeleteAllSeries(chartApi);
            ApplySeries(chartApi, context.Workbook, sheet, request.Param("series"));
            ApplyAxes(chartApi, request.Param("axes"));
            var exportPath = request.StringParam("exportPath");
            if (exportPath.Length > 0) chartApi.Export(Path.GetFullPath(exportPath), "PNG");
            return new { chart = Snapshot(chartObjectApi, Convert.ToString(((dynamic)sheet).Name) ?? string.Empty) };
        }
        catch (Exception exception)
        {
            throw new OfficeWorkerException("chart_operation_failed", exception.Message, null, exception);
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

    private static object Snapshot(dynamic chartObject, string sheetName)
    {
        object? chart = null;
        object? seriesCollection = null;
        try
        {
            chart = chartObject.Chart;
            dynamic chartApi = chart;
            seriesCollection = chartApi.SeriesCollection();
            dynamic seriesApi = seriesCollection;
            var series = new List<object>();
            for (var index = 1; index <= Convert.ToInt32(seriesApi.Count); index++)
            {
                object? item = null;
                try
                {
                    item = seriesApi.Item(index);
                    dynamic itemApi = item;
                    series.Add(new
                    {
                        index,
                        name = Safe(() => itemApi.Name),
                        formula = Safe(() => itemApi.Formula),
                        chartType = Safe(() => itemApi.ChartType),
                        axisGroup = Convert.ToInt32(Safe(() => itemApi.AxisGroup) ?? 1) == 2 ? "secondary" : "primary",
                    });
                }
                finally { ComInterop.Release(item); }
            }
            return new
            {
                sheet = sheetName,
                name = Convert.ToString(chartObject.Name),
                title = Convert.ToBoolean(chartApi.HasTitle) ? Safe(() => chartApi.ChartTitle.Text) : string.Empty,
                chartType = Convert.ToInt32(chartApi.ChartType),
                style = Safe(() => chartApi.ChartStyle),
                position = new { left = Convert.ToDouble(chartObject.Left), top = Convert.ToDouble(chartObject.Top), width = Convert.ToDouble(chartObject.Width), height = Convert.ToDouble(chartObject.Height) },
                series,
            };
        }
        finally { ComInterop.Release(seriesCollection); ComInterop.Release(chart); }
    }

    private static void ApplySeries(dynamic chart, dynamic workbook, dynamic defaultSheet, JsonElement configs)
    {
        if (configs.ValueKind != JsonValueKind.Array) return;
        foreach (var config in configs.EnumerateArray())
        {
            object? collection = null;
            object? series = null;
            try
            {
                collection = chart.SeriesCollection();
                dynamic collectionApi = collection;
                var command = String(config, "command", "update");
                if (command == "add") series = collectionApi.NewSeries();
                else if (config.TryGetProperty("index", out var index)) series = collectionApi.Item(index.GetInt32());
                else if (config.TryGetProperty("matchName", out var matchName)) series = FindSeries(collectionApi, matchName.GetString() ?? string.Empty);
                if (series is null) throw new OfficeWorkerException("series_not_found", "找不到要编辑的数据系列");
                dynamic seriesApi = series;
                if (command == "delete") { seriesApi.Delete(); continue; }
                if (config.TryGetProperty("formula", out var formula)) seriesApi.Formula = formula.GetString();
                if (config.TryGetProperty("name", out var name)) seriesApi.Name = name.GetString();
                SetSeriesValues(seriesApi, "Values", workbook, defaultSheet, config, "values");
                SetSeriesValues(seriesApi, "XValues", workbook, defaultSheet, config, config.TryGetProperty("categories", out _) ? "categories" : "xValues");
                if (config.TryGetProperty("chartType", out var chartType)) seriesApi.ChartType = ResolveChartType(chartType.GetString() ?? string.Empty);
                if (config.TryGetProperty("axisGroup", out var axisGroup)) seriesApi.AxisGroup = axisGroup.GetString() == "secondary" ? 2 : 1;
                if (config.TryGetProperty("smooth", out var smooth)) { try { seriesApi.Smooth = smooth.GetBoolean(); } catch { } }
                if (config.TryGetProperty("dataLabels", out var labels)) ApplyDataLabels(seriesApi, labels);
            }
            finally { ComInterop.Release(series); ComInterop.Release(collection); }
        }
    }

    private static void ApplyAxes(dynamic chart, JsonElement configs)
    {
        if (configs.ValueKind != JsonValueKind.Array) return;
        foreach (var config in configs.EnumerateArray())
        {
            object? axis = null;
            try
            {
                axis = chart.Axes(String(config, "kind") == "category" ? 1 : 2, String(config, "group") == "secondary" ? 2 : 1);
                dynamic axisApi = axis;
                if (config.TryGetProperty("title", out var title))
                {
                    axisApi.HasTitle = !string.IsNullOrWhiteSpace(title.GetString());
                    if (axisApi.HasTitle) axisApi.AxisTitle.Text = title.GetString();
                }
                if (config.TryGetProperty("minimum", out var minimum)) { axisApi.MinimumScaleIsAuto = false; axisApi.MinimumScale = minimum.GetDouble(); }
                if (config.TryGetProperty("maximum", out var maximum)) { axisApi.MaximumScaleIsAuto = false; axisApi.MaximumScale = maximum.GetDouble(); }
                if (config.TryGetProperty("majorUnit", out var majorUnit)) { axisApi.MajorUnitIsAuto = false; axisApi.MajorUnit = majorUnit.GetDouble(); }
                if (config.TryGetProperty("numberFormat", out var format)) axisApi.TickLabels.NumberFormat = format.GetString();
                if (config.TryGetProperty("reverse", out var reverse)) axisApi.ReversePlotOrder = reverse.GetBoolean();
            }
            finally { ComInterop.Release(axis); }
        }
    }

    private static void ApplyDataLabels(dynamic series, JsonElement config)
    {
        if (config.ValueKind != JsonValueKind.Object) return;
        if (config.TryGetProperty("enabled", out var enabled) && !enabled.GetBoolean())
        {
            try { series.DataLabels().Delete(); } catch { }
            return;
        }
        series.ApplyDataLabels();
        object? labels = null;
        try
        {
            labels = series.DataLabels();
            dynamic labelsApi = labels;
            if (config.TryGetProperty("showValue", out var value)) labelsApi.ShowValue = value.GetBoolean();
            if (config.TryGetProperty("showCategoryName", out var category)) labelsApi.ShowCategoryName = category.GetBoolean();
            if (config.TryGetProperty("showSeriesName", out var name)) labelsApi.ShowSeriesName = name.GetBoolean();
            if (config.TryGetProperty("numberFormat", out var format)) labelsApi.NumberFormat = format.GetString();
        }
        finally { ComInterop.Release(labels); }
    }

    private static void SetSeriesValues(dynamic series, string property, dynamic workbook, dynamic defaultSheet, JsonElement config, string name)
    {
        if (!config.TryGetProperty(name, out var value)) return;
        if (value.ValueKind == JsonValueKind.String)
        {
            object? range = null;
            try
            {
                range = ResolveRange(workbook, defaultSheet, value.GetString() ?? string.Empty);
                series.GetType().InvokeMember(property, System.Reflection.BindingFlags.SetProperty, null, series, new object?[] { range });
            }
            finally { ComInterop.Release(range); }
        }
        else
        {
            var values = value.ValueKind == JsonValueKind.Array ? value.EnumerateArray().Select(item => item.ToString()).ToArray() : [];
            series.GetType().InvokeMember(property, System.Reflection.BindingFlags.SetProperty, null, series, new object?[] { values });
        }
    }

    private static object ResolveRange(dynamic workbook, dynamic defaultSheet, string reference)
    {
        var parts = reference.Split('!', 2);
        if (parts.Length == 1) return defaultSheet.Range(reference);
        object? sheet = null;
        try
        {
            sheet = workbook.Worksheets.Item(parts[0].Trim('\''));
            return ((dynamic)sheet).Range(parts[1]);
        }
        finally { ComInterop.Release(sheet); }
    }

    private static object? FindSeries(dynamic collection, string name)
    {
        for (var index = 1; index <= Convert.ToInt32(collection.Count); index++)
        {
            object? series = collection.Item(index);
            if (string.Equals(Convert.ToString(((dynamic)series).Name), name, StringComparison.OrdinalIgnoreCase)) return series;
            ComInterop.Release(series);
        }
        return null;
    }

    private static void DeleteAllSeries(dynamic chart)
    {
        object? collection = null;
        try
        {
            collection = chart.SeriesCollection();
            dynamic collectionApi = collection;
            while (Convert.ToInt32(collectionApi.Count) > 0)
            {
                object? series = null;
                try { series = collectionApi.Item(1); ((dynamic)series).Delete(); }
                finally { ComInterop.Release(series); }
            }
        }
        finally { ComInterop.Release(collection); }
    }

    private static void SetPosition(dynamic chartObject, OfficeActionRequest request)
    {
        if (request.Param("left").ValueKind == JsonValueKind.Number) chartObject.Left = request.DoubleParam("left");
        if (request.Param("top").ValueKind == JsonValueKind.Number) chartObject.Top = request.DoubleParam("top");
        if (request.Param("width").ValueKind == JsonValueKind.Number) chartObject.Width = request.DoubleParam("width");
        if (request.Param("height").ValueKind == JsonValueKind.Number) chartObject.Height = request.DoubleParam("height");
    }

    private static string String(JsonElement value, string name, string fallback = "") => value.TryGetProperty(name, out var property) && property.ValueKind == JsonValueKind.String ? property.GetString() ?? fallback : fallback;
    private static object? Safe(Func<object?> value) { try { return value(); } catch { return null; } }
    private static int ResolveChartType(string name) => name.ToLowerInvariant() switch { "line" => 4, "linemarkers" => 65, "pie" => 5, "doughnut" => -4120, "bar" => 57, "area" => 1, "scatter" => -4169, "bubble" => 15, "radar" => -4151, _ => 51 };
}
