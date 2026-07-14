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
        var (chartTypeCode, chartTypeName) = ResolveChartType(request);
        object? chartObjects = null;
        object? chartObject = null;
        object? chart = null;
        object? seriesCollection = null;
        try
        {
            dynamic sheetApi = sheet;
            dynamic rangeApi = range;
            chartObjects = sheetApi.ChartObjects();
            dynamic chartObjectsApi = chartObjects;
            chartObject = chartObjectsApi.Add(
                Convert.ToDouble(rangeApi.Left) + Convert.ToDouble(rangeApi.Width) + 20,
                Convert.ToDouble(rangeApi.Top), 420, 260);
            if (chartObject is null)
                throw new OfficeWorkerException("chart_object_missing", $"ChartObjects.Add 返回 null；sheet={request.ExcelTarget().SheetName}");
            chart = ((dynamic)chartObject).Chart;
            if (chart is null)
                throw new OfficeWorkerException("chart_object_missing", $"新插入 chartObject 没有 Chart 成员");
            dynamic chartApi = chart;
            chartApi.SetSourceData(range);
            chartApi.ChartType = chartTypeCode;
            seriesCollection = chartApi.SeriesCollection();
            var verification = VerifyInsertedChart(chartObject, chart, seriesCollection);
            if (!verification.Ok)
                throw new OfficeWorkerException(
                    "chart_verification_failed",
                    $"图表对象已创建，但插入后校验失败：{verification.ChartName}",
                    verification);
            return Done(request, "chart", "已插入 Excel 图表",
                new
                {
                    chartId = verification.ChartId,
                    chartName = verification.ChartName,
                    sheetName = verification.SheetName,
                    chartType = chartTypeName,
                    seriesCount = verification.SeriesCount,
                    hasTitle = verification.HasTitle,
                    position = verification.Position,
                    size = verification.Size,
                    verification = new
                    {
                        ok = verification.Ok,
                        chartObjectExists = verification.ChartObjectExists,
                        chartNonNull = verification.ChartNonNull,
                        chartVisible = verification.ChartVisible,
                        sheetVisible = verification.SheetVisible,
                        seriesCount = verification.SeriesCount,
                        hasTitle = verification.HasTitle,
                        titleText = verification.TitleText,
                        anchorReadable = verification.AnchorReadable,
                        topLeftCell = verification.TopLeftCell,
                        bottomRightCell = verification.BottomRightCell,
                        widthPx = verification.WidthPx,
                        heightPx = verification.HeightPx,
                        checks = verification.Checks,
                    }
                });
        }
        catch (OfficeWorkerException) { throw; }
        catch (Exception ex)
        {
            throw new OfficeWorkerException("chart_insert_failed", $"插入图表失败：{ex.Message}", null, ex);
        }
        finally
        {
            ComInterop.Release(seriesCollection);
            ComInterop.Release(chart);
            ComInterop.Release(chartObject);
            ComInterop.Release(chartObjects);
            ComInterop.Release(range);
            ComInterop.Release(sheet);
        }
    }

    private static (int code, string name) ResolveChartType(OfficeActionRequest request)
    {
        var type = request.StringParam("chartType", "column");
        return ChartType(type) switch
        {
            4 => (4, "line"),
            5 => (5, "pie"),
            57 => (57, "bar"),
            1 => (1, "area"),
            -4169 => (-4169, "scatter"),
            51 => (51, "column"),
            var unknown => throw new OfficeWorkerException(
                "unsupported_chart_type",
                $"未识别的 chartType={request.StringParam("chartType", "column")}；请使用 line/pie/bar/area/scatter/column"),
        };
    }

    internal sealed record ChartPostWrite(
        bool Ok,
        bool ChartObjectExists,
        bool ChartNonNull,
        bool ChartVisible,
        bool SheetVisible,
        string ChartId,
        string ChartName,
        string SheetName,
        int SeriesCount,
        bool HasTitle,
        string TitleText,
        bool AnchorReadable,
        string TopLeftCell,
        string BottomRightCell,
        double WidthPx,
        double HeightPx,
        Position Position,
        Size Size,
        IReadOnlyList<object> Checks);

    internal sealed record Position(double Left, double Top);
    internal sealed record Size(double Width, double Height);

    internal static ChartPostWrite VerifyInsertedChart(object chartObject, object chart, object seriesCollection)
    {
        var checks = new List<object>();
        checks.Add(new { name = "chart-object-exists", ok = true, message = "ChartObject 已创建" });
        checks.Add(new { name = "chart-non-null", ok = true, message = "Chart 成员有效" });

        string chartName = string.Empty;
        string sheetName = string.Empty;
        string topLeftCell = string.Empty;
        string bottomRightCell = string.Empty;
        bool chartVisible = false;
        bool sheetVisible = false;
        double left = 0, top = 0, width = 0, height = 0;
        object? parent = null;
        object? topLeft = null;
        object? bottomRight = null;
        try
        {
            dynamic obj = chartObject!;
            chartName = Convert.ToString(obj.Name) ?? string.Empty;
            chartVisible = Convert.ToBoolean(obj.Visible);
            parent = obj.Parent;
            sheetName = Convert.ToString(((dynamic)parent).Name) ?? string.Empty;
            sheetVisible = Convert.ToInt32(((dynamic)parent).Visible) == -1;
            left = Convert.ToDouble(obj.Left);
            top = Convert.ToDouble(obj.Top);
            width = Convert.ToDouble(obj.Width);
            height = Convert.ToDouble(obj.Height);
            topLeft = obj.TopLeftCell;
            bottomRight = obj.BottomRightCell;
            topLeftCell = Convert.ToString(((dynamic)topLeft).Address) ?? string.Empty;
            bottomRightCell = Convert.ToString(((dynamic)bottomRight).Address) ?? string.Empty;
        }
        catch (Exception ex)
        {
            checks.Add(new { name = "chart-geometry-readable", ok = false, message = ex.Message });
        }
        finally
        {
            ComInterop.Release(bottomRight);
            ComInterop.Release(topLeft);
            ComInterop.Release(parent);
        }
        checks.Add(new { name = "chart-name-present", ok = !string.IsNullOrWhiteSpace(chartName), message = $"Name={chartName}" });
        checks.Add(new { name = "sheet-name-present", ok = !string.IsNullOrWhiteSpace(sheetName), message = $"Sheet={sheetName}" });
        checks.Add(new { name = "chart-size-positive", ok = width > 0 && height > 0, message = $"size={width}x{height}" });
        checks.Add(new { name = "chart-visible", ok = chartVisible, message = $"ChartObject.Visible={chartVisible}" });
        checks.Add(new { name = "sheet-visible", ok = sheetVisible, message = $"Worksheet.Visible={sheetVisible}" });
        var anchorReadable = !string.IsNullOrWhiteSpace(topLeftCell) && !string.IsNullOrWhiteSpace(bottomRightCell);
        checks.Add(new { name = "chart-anchor-readable", ok = anchorReadable, message = $"anchor={topLeftCell}:{bottomRightCell}" });

        int seriesCount = 0;
        try
        {
            seriesCount = Convert.ToInt32(((dynamic)seriesCollection).Count);
            checks.Add(new { name = "chart-series-readable", ok = true, message = $"seriesCount={seriesCount}" });
        }
        catch (Exception ex)
        {
            checks.Add(new { name = "chart-series-readable", ok = false, message = ex.Message });
        }
        checks.Add(new { name = "chart-has-series", ok = seriesCount > 0, message = seriesCount > 0 ? "包含至少一个系列" : "无任何系列，图表将不可见" });

        bool hasTitle = false;
        string titleText = string.Empty;
        object? chartTitle = null;
        try
        {
            dynamic chartApi = chart;
            hasTitle = Convert.ToBoolean(chartApi.HasTitle);
            if (hasTitle)
            {
                try
                {
                    chartTitle = chartApi.ChartTitle;
                    titleText = Convert.ToString(((dynamic)chartTitle).Text) ?? string.Empty;
                }
                catch { /* 非关键 */ }
            }
            checks.Add(new { name = "chart-title", ok = true, message = $"HasTitle={hasTitle} text='{titleText}'" });
        }
        catch (Exception ex)
        {
            checks.Add(new { name = "chart-title", ok = false, message = ex.Message });
        }
        finally { ComInterop.Release(chartTitle); }

        var ok = !string.IsNullOrWhiteSpace(chartName) && !string.IsNullOrWhiteSpace(sheetName) &&
            chartVisible && sheetVisible && seriesCount > 0 && width > 0 && height > 0 && anchorReadable;
        return new ChartPostWrite(
            Ok: ok,
            ChartObjectExists: true,
            ChartNonNull: true,
            ChartVisible: chartVisible,
            SheetVisible: sheetVisible,
            ChartId: $"{sheetName}!{chartName}",
            ChartName: chartName,
            SheetName: sheetName,
            SeriesCount: seriesCount,
            HasTitle: hasTitle,
            TitleText: titleText,
            AnchorReadable: anchorReadable,
            TopLeftCell: topLeftCell,
            BottomRightCell: bottomRightCell,
            WidthPx: width,
            HeightPx: height,
            Position: new Position(left, top),
            Size: new Size(width, height),
            Checks: checks);
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
            var name = request.StringParam("name", $"AI_Pivot_{DateTime.Now:yyyyMMddHHmmssfff}");
            var destination = request.StringParam("destination");
            var destinationPlan = ParsePivotDestination(destination);
            string destinationAddress;
            if (destinationPlan.UseDedicatedSheet)
            {
                destinationSheet = EnsurePivotSheet(context);
                destinationAddress = NextPivotDestinationAddress(destinationSheet);
                destination = $"{Convert.ToString(((dynamic)destinationSheet).Name)}!{destinationAddress}";
            }
            else
            {
                destinationSheet = destinationPlan.SheetName is not null
                    ? context.Workbook.Worksheets.Item(destinationPlan.SheetName)
                    : sheet;
                destinationAddress = destinationPlan.Address;
            }
            dynamic destinationSheetApi = destinationSheet;
            destinationRange = destinationSheetApi.Range(destinationAddress);
            dynamic rangeApi = range;
            var sourceAddress = Convert.ToString(rangeApi.Address[true, true, 1, true])
                ?? throw new OfficeWorkerException("pivot_source_invalid", "无法读取透视表源区域地址");
            caches = context.Workbook.PivotCaches();
            cache = ((dynamic)caches).Create(1, sourceAddress);
            pivot = ((dynamic)cache).CreatePivotTable(destinationRange, name);
            if (pivot is null) throw new OfficeWorkerException("pivot_object_missing", "CreatePivotTable 未返回透视表对象");
            SetPivotOrientations(pivot, request.Param("rowFields"), 1, "rowFields");
            SetPivotOrientations(pivot, request.Param("columnFields"), 2, "columnFields");
            SetPivotOrientations(pivot, request.Param("filterFields"), 3, "filterFields");
            AddPivotDataFields(pivot, request.Param("dataFields"));

            var readback = VerifyPivotTable(pivot, name, sourceAddress);
            return Done(request, "pivot-table", "已创建数据透视表", new { destination, readback });
        }
        catch (OfficeWorkerException) { throw; }
        catch (Exception ex)
        {
            throw new OfficeWorkerException("pivot_create_failed", $"创建数据透视表失败：{ex.Message}", null, ex);
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

    private static object EnsurePivotSheet(ExcelActionContext context)
    {
        object? sheets = null;
        try
        {
            sheets = context.Workbook.Worksheets;
            dynamic sheetsApi = sheets;
            for (var i = 1; i <= Convert.ToInt32(sheetsApi.Count); i++)
            {
                object existingSheet = sheetsApi.Item(i);
                if (string.Equals(Convert.ToString(((dynamic)existingSheet).Name), "Pivots", StringComparison.OrdinalIgnoreCase))
                    return existingSheet;
                ComInterop.Release(existingSheet);
            }
            object? afterSheet = null;
            object? newSheet;
            try
            {
                afterSheet = sheetsApi.Item(sheetsApi.Count);
                newSheet = sheetsApi.Add(Type.Missing, afterSheet);
            }
            finally { ComInterop.Release(afterSheet); }
            dynamic newSheetApi = newSheet;
            newSheetApi.Name = "Pivots";
            return newSheet;
        }
        finally
        {
            ComInterop.Release(sheets);
        }
    }

    internal sealed record PivotDestinationPlan(bool UseDedicatedSheet, string? SheetName, string Address);

    internal static PivotDestinationPlan ParsePivotDestination(string? destination)
    {
        if (string.IsNullOrWhiteSpace(destination)) return new PivotDestinationPlan(true, null, "A1");
        var parts = destination.Split('!', 2);
        return parts.Length == 2
            ? new PivotDestinationPlan(false, parts[0].Trim('\''), parts[1])
            : new PivotDestinationPlan(false, null, parts[0]);
    }

    private static string NextPivotDestinationAddress(object sheet)
    {
        object? pivots = null;
        var lastRow = 0;
        try
        {
            pivots = ((dynamic)sheet).PivotTables();
            dynamic pivotsApi = pivots;
            for (var index = 1; index <= Convert.ToInt32(pivotsApi.Count); index++)
            {
                object? pivot = null;
                object? tableRange = null;
                object? rows = null;
                try
                {
                    pivot = pivotsApi.Item(index);
                    tableRange = ((dynamic)pivot).TableRange2;
                    rows = ((dynamic)tableRange).Rows;
                    var bottom = Convert.ToInt32(((dynamic)tableRange).Row) + Convert.ToInt32(((dynamic)rows).Count) - 1;
                    lastRow = Math.Max(lastRow, bottom);
                }
                finally
                {
                    ComInterop.Release(rows);
                    ComInterop.Release(tableRange);
                    ComInterop.Release(pivot);
                }
            }
        }
        finally { ComInterop.Release(pivots); }
        return lastRow == 0 ? "A1" : $"A{lastRow + 3}";
    }

    internal sealed record PivotPostWrite(
        string PivotName,
        string SourceAddress,
        string? DestinationRange,
        int RowFieldCount,
        int ColumnFieldCount,
        int FilterFieldCount,
        int DataFieldCount,
        string? TableRange1,
        string? TableRange2,
        string? DataBodyRange,
        object Verification);

    internal static PivotPostWrite VerifyPivotTable(object pivot, string name, string sourceAddress)
    {
        var checks = new List<object>();
        int rowFieldCount = 0, columnFieldCount = 0, filterFieldCount = 0, dataFieldCount = 0;
        string? tableRange1 = null, tableRange2 = null, dataBodyRange = null;
        var fieldOrientationsReadable = true;
        var cacheReadable = false;
        dynamic pivotApi = pivot;

        try { rowFieldCount = ReadPivotCollectionCount(pivot, "row"); checks.Add(new { name = "pivot-rowfields-readable", ok = true, message = $"rowFields={rowFieldCount}" }); }
        catch (Exception ex) { fieldOrientationsReadable = false; checks.Add(new { name = "pivot-rowfields-readable", ok = false, message = ex.Message }); }
        try { columnFieldCount = ReadPivotCollectionCount(pivot, "column"); checks.Add(new { name = "pivot-columnfields-readable", ok = true, message = $"columnFields={columnFieldCount}" }); }
        catch (Exception ex) { fieldOrientationsReadable = false; checks.Add(new { name = "pivot-columnfields-readable", ok = false, message = ex.Message }); }
        try { filterFieldCount = ReadPivotCollectionCount(pivot, "filter"); checks.Add(new { name = "pivot-filterfields-readable", ok = true, message = $"filterFields={filterFieldCount}" }); }
        catch (Exception ex) { fieldOrientationsReadable = false; checks.Add(new { name = "pivot-filterfields-readable", ok = false, message = ex.Message }); }
        try { dataFieldCount = ReadPivotCollectionCount(pivot, "data"); checks.Add(new { name = "pivot-datafields-readable", ok = true, message = $"dataFields={dataFieldCount}" }); }
        catch (Exception ex) { fieldOrientationsReadable = false; checks.Add(new { name = "pivot-datafields-readable", ok = false, message = ex.Message }); }

        object? pivotCache = null;
        try { pivotCache = pivotApi.PivotCache(); cacheReadable = pivotCache is not null; checks.Add(new { name = "pivot-cache-readable", ok = cacheReadable, message = cacheReadable ? "缓存可读" : "缓存为空" }); }
        catch (Exception ex) { checks.Add(new { name = "pivot-cache-readable", ok = false, message = ex.Message }); }
        finally { ComInterop.Release(pivotCache); }

        tableRange1 = ReadPivotRangeAddress(pivot, "table1", checks);
        tableRange2 = ReadPivotRangeAddress(pivot, "table2", checks);
        dataBodyRange = ReadPivotRangeAddress(pivot, "data", checks, optional: true);
        var fieldCount = rowFieldCount + columnFieldCount + filterFieldCount + dataFieldCount;
        var ok = cacheReadable && fieldOrientationsReadable && tableRange1 is not null && fieldCount > 0;
        var verification = new
        {
            ok,
            objectExists = true,
            cacheReadable,
            fieldOrientationsReadable,
            destinationReadable = tableRange1 is not null,
            checks,
        };
        var readback = new PivotPostWrite(name, sourceAddress, tableRange1, rowFieldCount, columnFieldCount,
            filterFieldCount, dataFieldCount, tableRange1, tableRange2, dataBodyRange, verification);
        if (!ok) throw new OfficeWorkerException("pivot_verification_failed", $"透视表创建后读回校验失败：{name}", readback);
        return readback;
    }

    private static int ReadPivotCollectionCount(object pivot, string kind)
    {
        object? fields = null;
        try
        {
            dynamic api = pivot;
            fields = kind switch
            {
                "row" => api.RowFields,
                "column" => api.ColumnFields,
                "filter" => api.PageFields,
                _ => api.DataFields,
            };
            return fields is null ? 0 : Convert.ToInt32(((dynamic)fields).Count);
        }
        finally { ComInterop.Release(fields); }
    }

    private static string? ReadPivotRangeAddress(object pivot, string kind, List<object> checks, bool optional = false)
    {
        object? range = null;
        try
        {
            dynamic api = pivot;
            range = kind switch
            {
                "table1" => api.TableRange1,
                "table2" => api.TableRange2,
                _ => api.DataBodyRange,
            };
            var address = range is null ? null : Convert.ToString(((dynamic)range).Address);
            checks.Add(new { name = $"pivot-{kind}-readable", ok = optional || address is not null, message = $"range={address ?? "(空)"}" });
            return address;
        }
        catch (Exception ex)
        {
            checks.Add(new { name = $"pivot-{kind}-readable", ok = optional, message = ex.Message });
            return null;
        }
        finally { ComInterop.Release(range); }
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

    private static void SetPivotOrientations(object pivot, System.Text.Json.JsonElement fields, int orientation, string parameterName)
    {
        if (fields.ValueKind != System.Text.Json.JsonValueKind.Array) return;
        foreach (var field in fields.EnumerateArray())
        {
            var name = field.GetString();
            if (string.IsNullOrWhiteSpace(name)) continue;
            object? pivotField = null;
            try
            {
                pivotField = ((dynamic)pivot).PivotFields(name);
                if (pivotField is null) throw new OfficeWorkerException("pivot_field_not_found", $"{parameterName} 字段不存在：{name}");
                ((dynamic)pivotField).Orientation = orientation;
            }
            catch (Exception ex)
            {
                throw new OfficeWorkerException(
                    "pivot_field_not_found",
                    $"{parameterName} 字段不可用：{name}；{ex.Message}",
                    null,
                    ex);
            }
            finally { ComInterop.Release(pivotField); }
        }
    }

    private static void AddPivotDataFields(object pivot, System.Text.Json.JsonElement fields)
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
            try
            {
                pivotField = ((dynamic)pivot).PivotFields(name);
                if (pivotField is null) throw new OfficeWorkerException("pivot_field_not_found", $"dataFields 字段不存在：{name}");
                ((dynamic)pivot).AddDataField(pivotField, caption, aggregate);
            }
            catch (Exception ex)
            {
                throw new OfficeWorkerException(
                    "pivot_field_not_found",
                    $"AddDataField 失败（field='{name}', function='{function}'）：{ex.Message}",
                    null,
                    ex);
            }
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
