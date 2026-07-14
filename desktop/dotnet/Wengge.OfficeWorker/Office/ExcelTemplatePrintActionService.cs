using System.Text.Json;
using Wengge.OfficeWorker.Com;

namespace Wengge.OfficeWorker.Office;

internal sealed class ExcelTemplatePrintActionService(OfficeApplicationProvider applications)
{
    public object Execute(OfficeActionRequest request)
    {
        using var context = new ExcelActionContext(applications, request);
        if (request.Operation is "captureWorkbookTemplate" or "inspectWorkbookFormatting")
            return OfficeActionResults.Done(request, "com", "已检查工作簿格式", Capture(context), []);
        if (request.Operation == "inspectPrintSettings")
            return OfficeActionResults.Done(request, "com", "已检查打印设置", InspectPrint(context, request), []);
        var data = request.Operation == "configurePrint" ? ConfigurePrint(context, request) : ApplyTemplate(context, request);
        context.Save(request);
        return OfficeActionResults.Done(request, "com", request.Operation == "configurePrint" ? "已配置打印设置" : "已应用专业工作簿模板", data,
            [new OfficeChange(request.Operation == "configurePrint" ? "print-settings" : "workbook-template", request.Target, request.Operation == "configurePrint" ? "已配置打印设置" : "已应用工作簿模板")]);
    }

    private static object Capture(ExcelActionContext context)
    {
        var sheetRules = new List<object>();
        object? sheets = null;
        try
        {
            sheets = context.Workbook.Worksheets;
            dynamic sheetsApi = sheets;
            for (var index = 1; index <= Convert.ToInt32(sheetsApi.Count); index++)
            {
                object? sheet = null;
                object? used = null;
                object? rows = null;
                object? header = null;
                object? font = null;
                object? headerFont = null;
                object? headerInterior = null;
                object? pageSetup = null;
                try
                {
                    sheet = sheetsApi.Item(index); dynamic sheetApi = sheet;
                    used = sheetApi.UsedRange; dynamic usedApi = used;
                    rows = usedApi.Rows;
                    header = ((dynamic)rows).Item(1);
                    font = usedApi.Font;
                    headerFont = ((dynamic)header).Font;
                    headerInterior = ((dynamic)header).Interior;
                    pageSetup = sheetApi.PageSetup;
                    dynamic setup = pageSetup;
                    sheetRules.Add(new
                    {
                        name = Convert.ToString(sheetApi.Name),
                        usedRange = Safe(() => usedApi.Address[false, false]),
                        rows = Convert.ToInt32(usedApi.Rows.Count),
                        columns = Convert.ToInt32(usedApi.Columns.Count),
                        baseStyle = new { fontName = Safe(() => ((dynamic)font).Name), fontSize = Safe(() => ((dynamic)font).Size), fontColor = Safe(() => ((dynamic)font).Color) },
                        headerStyle = new { fillColor = Safe(() => ((dynamic)headerInterior).Color), fontColor = Safe(() => ((dynamic)headerFont).Color), bold = Safe(() => ((dynamic)headerFont).Bold), rowHeight = Safe(() => ((dynamic)header).RowHeight) },
                        print = PrintSnapshot(setup),
                    });
                }
                finally
                {
                    ComInterop.Release(pageSetup); ComInterop.Release(headerInterior); ComInterop.Release(headerFont);
                    ComInterop.Release(font); ComInterop.Release(header); ComInterop.Release(rows); ComInterop.Release(used); ComInterop.Release(sheet);
                }
            }
            return new
            {
                template = new
                {
                    version = 1,
                    capturedFrom = Convert.ToString(context.Workbook.Name),
                    capturedAt = DateTimeOffset.UtcNow,
                    sheets = sheetRules,
                },
                sheetCount = sheetRules.Count,
            };
        }
        finally { ComInterop.Release(sheets); }
    }

    private static object ApplyTemplate(ExcelActionContext context, OfficeActionRequest request)
    {
        var preset = request.StringParam("preset", "professional");
        var accent = preset switch { "financial" => "217346", "dashboard" => "202124", "minimal" => "5F6368", _ => "1F4E79" };
        var headerFill = preset == "minimal" ? "E8EAED" : accent;
        var headerFontColor = preset == "minimal" ? "202124" : "FFFFFF";
        var requestedSheets = request.Param("sheetNames").ValueKind == JsonValueKind.Array
            ? request.Param("sheetNames").EnumerateArray().Select(value => value.GetString()).Where(value => !string.IsNullOrWhiteSpace(value)).ToHashSet(StringComparer.OrdinalIgnoreCase)
            : [];
        var applied = new List<object>();
        object? sheets = null;
        try
        {
            sheets = context.Workbook.Worksheets;
            dynamic sheetsApi = sheets;
            for (var index = 1; index <= Convert.ToInt32(sheetsApi.Count); index++)
            {
                object? sheet = null;
                object? used = null;
                object? header = null;
                object? font = null;
                object? headerFont = null;
                object? interior = null;
                object? columns = null;
                object? rows = null;
                try
                {
                    sheet = sheetsApi.Item(index); dynamic sheetApi = sheet;
                    var name = Convert.ToString(sheetApi.Name) ?? string.Empty;
                    if (requestedSheets.Count > 0 && !requestedSheets.Contains(name)) continue;
                    if (!request.BoolParam("allSheets", true) && !name.Equals(request.ExcelTarget().SheetName, StringComparison.OrdinalIgnoreCase)) continue;
                    used = sheetApi.UsedRange; dynamic usedApi = used;
                    if (Convert.ToInt32(usedApi.Cells.Count) == 1 && string.IsNullOrWhiteSpace(Convert.ToString(usedApi.Text))) continue;
                    font = usedApi.Font; dynamic fontApi = font;
                    fontApi.Name = request.StringParam("fontName", "微软雅黑");
                    fontApi.Size = request.DoubleParam("fontSize", 10.5);
                    header = usedApi.Rows.Item(1); dynamic headerApi = header;
                    headerFont = headerApi.Font; dynamic headerFontApi = headerFont;
                    headerFontApi.Bold = true; headerFontApi.Color = ExcelActionService.OleColor(headerFontColor);
                    interior = headerApi.Interior; ((dynamic)interior).Color = ExcelActionService.OleColor(headerFill);
                    headerApi.HorizontalAlignment = -4108; headerApi.WrapText = true; headerApi.RowHeight = 24;
                    columns = usedApi.Columns; rows = usedApi.Rows;
                    if (request.BoolParam("autoFit", true)) { ((dynamic)columns).AutoFit(); ((dynamic)rows).AutoFit(); }
                    sheetApi.Activate();
                    object? window = null;
                    try
                    {
                        window = context.App.ActiveWindow; dynamic windowApi = window;
                        windowApi.DisplayGridlines = request.BoolParam("showGridlines");
                        windowApi.FreezePanes = false;
                        var freezeRows = request.IntParam("freezeRows", 1);
                        if (freezeRows > 0)
                        {
                            object? cell = null;
                            try { cell = sheetApi.Cells.Item(freezeRows + 1, 1); ((dynamic)cell).Select(); windowApi.FreezePanes = true; }
                            finally { ComInterop.Release(cell); }
                        }
                    }
                    finally { ComInterop.Release(window); }
                    applied.Add(new { name, range = Safe(() => usedApi.Address[false, false]), rows = Convert.ToInt32(usedApi.Rows.Count), columns = Convert.ToInt32(usedApi.Columns.Count) });
                }
                finally
                {
                    ComInterop.Release(rows); ComInterop.Release(columns); ComInterop.Release(interior); ComInterop.Release(headerFont);
                    ComInterop.Release(font); ComInterop.Release(header); ComInterop.Release(used); ComInterop.Release(sheet);
                }
            }
            return new { preset, appliedSheets = applied, appliedSheetCount = applied.Count };
        }
        finally { ComInterop.Release(sheets); }
    }

    private static object InspectPrint(ExcelActionContext context, OfficeActionRequest request)
    {
        var settings = new List<object>();
        var targetSheets = RequestedSheetNames(request);
        object? sheets = null;
        try
        {
            sheets = context.Workbook.Worksheets; dynamic sheetsApi = sheets;
            for (var index = 1; index <= Convert.ToInt32(sheetsApi.Count); index++)
            {
                object? sheet = null;
                object? setup = null;
                try
                {
                    sheet = sheetsApi.Item(index); dynamic sheetApi = sheet;
                    var name = Convert.ToString(sheetApi.Name) ?? string.Empty;
                    if (targetSheets.Count > 0 && !targetSheets.Contains(name)) continue;
                    if (request.StringParam("sheetName").Length > 0 && !name.Equals(request.StringParam("sheetName"), StringComparison.OrdinalIgnoreCase)) continue;
                    settings.Add(WorksheetPrintSnapshot(sheetApi));
                }
                finally { ComInterop.Release(setup); ComInterop.Release(sheet); }
            }
            return new { progId = context.ProgId, settings, sheetCount = settings.Count };
        }
        finally { ComInterop.Release(sheets); }
    }

    private static object ConfigurePrint(ExcelActionContext context, OfficeActionRequest request)
    {
        var targetSheets = RequestedSheetNames(request);
        var settings = new List<object>();
        object? sheets = null;
        try
        {
            sheets = context.Workbook.Worksheets; dynamic sheetsApi = sheets;
            for (var index = 1; index <= Convert.ToInt32(sheetsApi.Count); index++)
            {
                object? sheet = null;
                object? setup = null;
                try
                {
                    sheet = sheetsApi.Item(index); dynamic sheetApi = sheet;
                    var name = Convert.ToString(sheetApi.Name) ?? string.Empty;
                    if (targetSheets.Count > 0 && !targetSheets.Contains(name)) continue;
                    if (targetSheets.Count == 0 && request.StringParam("sheetName").Length > 0 && !name.Equals(request.StringParam("sheetName"), StringComparison.OrdinalIgnoreCase)) continue;
                    setup = sheetApi.PageSetup; dynamic setupApi = setup;
                    if (request.StringParam("orientation").Length > 0) setupApi.Orientation = request.StringParam("orientation") == "landscape" ? 2 : 1;
                    if (request.StringParam("paperSize").Length > 0) setupApi.PaperSize = PaperSize(request.StringParam("paperSize"));
                    if (request.StringParam("printArea").Length > 0) setupApi.PrintArea = request.StringParam("printArea");
                    if (request.StringParam("repeatRows").Length > 0) setupApi.PrintTitleRows = request.StringParam("repeatRows");
                    if (request.StringParam("repeatColumns").Length > 0) setupApi.PrintTitleColumns = request.StringParam("repeatColumns");
                    ApplyMargins(setupApi, request);
                    if (request.Param("fitToOnePageWide").ValueKind == JsonValueKind.True)
                    {
                        setupApi.Zoom = false;
                        setupApi.FitToPagesWide = 1;
                        setupApi.FitToPagesTall = request.BoolParam("fitToOnePageTall") ? (object)1 : false;
                    }
                    else if (request.Param("scale").ValueKind == JsonValueKind.Number)
                    {
                        setupApi.Zoom = Math.Clamp(request.IntParam("scale"), 10, 400);
                    }
                    else
                    {
                        if (request.Param("fitToPagesWide").ValueKind == JsonValueKind.Number) { setupApi.Zoom = false; setupApi.FitToPagesWide = request.IntParam("fitToPagesWide"); }
                        if (request.Param("fitToPagesTall").ValueKind == JsonValueKind.Number) { setupApi.Zoom = false; setupApi.FitToPagesTall = request.IntParam("fitToPagesTall"); }
                    }
                    if (request.Param("marginLeft").ValueKind == JsonValueKind.Number) setupApi.LeftMargin = request.DoubleParam("marginLeft");
                    if (request.Param("marginRight").ValueKind == JsonValueKind.Number) setupApi.RightMargin = request.DoubleParam("marginRight");
                    if (request.Param("marginTop").ValueKind == JsonValueKind.Number) setupApi.TopMargin = request.DoubleParam("marginTop");
                    if (request.Param("marginBottom").ValueKind == JsonValueKind.Number) setupApi.BottomMargin = request.DoubleParam("marginBottom");
                    if (request.Param("centerHorizontally").ValueKind is JsonValueKind.True or JsonValueKind.False) setupApi.CenterHorizontally = request.BoolParam("centerHorizontally");
                    if (request.Param("centerVertically").ValueKind is JsonValueKind.True or JsonValueKind.False) setupApi.CenterVertically = request.BoolParam("centerVertically");
                    if (request.Param("printGridlines").ValueKind is JsonValueKind.True or JsonValueKind.False) setupApi.PrintGridlines = request.BoolParam("printGridlines");
                    if (request.Param("printHeadings").ValueKind is JsonValueKind.True or JsonValueKind.False) setupApi.PrintHeadings = request.BoolParam("printHeadings");
                    if (request.Param("blackAndWhite").ValueKind is JsonValueKind.True or JsonValueKind.False) setupApi.BlackAndWhite = request.BoolParam("blackAndWhite");
                    if (request.Param("draft").ValueKind is JsonValueKind.True or JsonValueKind.False) setupApi.Draft = request.BoolParam("draft");
                    if (request.StringParam("pageOrder").Length > 0) setupApi.Order = request.StringParam("pageOrder") == "overThenDown" ? 2 : 1;
                    if (request.Param("firstPageNumber").ValueKind == JsonValueKind.Number) setupApi.FirstPageNumber = request.IntParam("firstPageNumber");
                    if (request.StringParam("header").Length > 0) setupApi.CenterHeader = request.StringParam("header");
                    if (request.StringParam("footer").Length > 0) setupApi.CenterFooter = request.StringParam("footer");
                    ApplyHeaderFooter(setupApi, request.Param("headers"), header: true);
                    ApplyHeaderFooter(setupApi, request.Param("footers"), header: false);
                    if (request.BoolParam("clearPageBreaks")) sheetApi.ResetAllPageBreaks();
                    AddPageBreaks(sheetApi, request.Param("horizontalPageBreaks"), horizontal: true);
                    AddPageBreaks(sheetApi, request.Param("verticalPageBreaks"), horizontal: false);
                    settings.Add(WorksheetPrintSnapshot(sheetApi));
                }
                finally { ComInterop.Release(setup); ComInterop.Release(sheet); }
            }
            return new { settings, sheetCount = settings.Count };
        }
        finally { ComInterop.Release(sheets); }
    }

    private static HashSet<string> RequestedSheetNames(OfficeActionRequest request)
    {
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var requested = request.Param("sheetNames");
        if (requested.ValueKind != JsonValueKind.Array) return names;
        foreach (var value in requested.EnumerateArray())
        {
            if (value.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(value.GetString())) names.Add(value.GetString()!);
        }
        return names;
    }

    private static void ApplyMargins(dynamic setup, OfficeActionRequest request)
    {
        var margins = request.Param("margins");
        if (margins.ValueKind != JsonValueKind.Object) return;
        var unit = request.StringParam("marginUnit", "centimeters");
        if (TryNumber(margins, "top", out var top)) setup.TopMargin = PrintLength(top, unit);
        if (TryNumber(margins, "bottom", out var bottom)) setup.BottomMargin = PrintLength(bottom, unit);
        if (TryNumber(margins, "left", out var left)) setup.LeftMargin = PrintLength(left, unit);
        if (TryNumber(margins, "right", out var right)) setup.RightMargin = PrintLength(right, unit);
        if (TryNumber(margins, "header", out var header)) setup.HeaderMargin = PrintLength(header, unit);
        if (TryNumber(margins, "footer", out var footer)) setup.FooterMargin = PrintLength(footer, unit);
    }

    private static double PrintLength(double value, string unit) => unit switch
    {
        "points" => value,
        "inches" => value * 72d,
        _ => value * 72d / 2.54d,
    };

    private static void ApplyHeaderFooter(dynamic setup, JsonElement values, bool header)
    {
        if (values.ValueKind != JsonValueKind.Object) return;
        if (TryString(values, "left", out var left))
        {
            if (header) setup.LeftHeader = left; else setup.LeftFooter = left;
        }
        if (TryString(values, "center", out var center))
        {
            if (header) setup.CenterHeader = center; else setup.CenterFooter = center;
        }
        if (TryString(values, "right", out var right))
        {
            if (header) setup.RightHeader = right; else setup.RightFooter = right;
        }
    }

    private static void AddPageBreaks(dynamic sheet, JsonElement addresses, bool horizontal)
    {
        if (addresses.ValueKind != JsonValueKind.Array) return;
        object? breaks = null;
        try
        {
            breaks = horizontal ? sheet.HPageBreaks : sheet.VPageBreaks;
            dynamic breaksApi = breaks;
            foreach (var address in addresses.EnumerateArray())
            {
                if (address.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(address.GetString())) continue;
                object? range = null;
                object? pageBreak = null;
                try
                {
                    range = sheet.Range(address.GetString());
                    pageBreak = breaksApi.Add(range);
                }
                finally
                {
                    ComInterop.Release(pageBreak);
                    ComInterop.Release(range);
                }
            }
        }
        finally
        {
            ComInterop.Release(breaks);
        }
    }

    private static object WorksheetPrintSnapshot(dynamic sheet)
    {
        object? setup = null;
        try
        {
            setup = sheet.PageSetup;
            dynamic setupApi = setup;
            return new
            {
                sheet = Convert.ToString(sheet.Name),
                printArea = Safe(() => setupApi.PrintArea),
                orientation = Convert.ToInt32(Safe(() => setupApi.Orientation) ?? 1) == 2 ? "landscape" : "portrait",
                paperSize = Safe(() => setupApi.PaperSize),
                marginsPoints = new
                {
                    top = Safe(() => setupApi.TopMargin),
                    bottom = Safe(() => setupApi.BottomMargin),
                    left = Safe(() => setupApi.LeftMargin),
                    right = Safe(() => setupApi.RightMargin),
                    header = Safe(() => setupApi.HeaderMargin),
                    footer = Safe(() => setupApi.FooterMargin),
                },
                repeatRows = Safe(() => setupApi.PrintTitleRows),
                repeatColumns = Safe(() => setupApi.PrintTitleColumns),
                zoom = Safe(() => setupApi.Zoom),
                fitToPagesWide = Safe(() => setupApi.FitToPagesWide),
                fitToPagesTall = Safe(() => setupApi.FitToPagesTall),
                centerHorizontally = Safe(() => setupApi.CenterHorizontally),
                centerVertically = Safe(() => setupApi.CenterVertically),
                printGridlines = Safe(() => setupApi.PrintGridlines),
                printHeadings = Safe(() => setupApi.PrintHeadings),
                headers = new { left = Safe(() => setupApi.LeftHeader), center = Safe(() => setupApi.CenterHeader), right = Safe(() => setupApi.RightHeader) },
                footers = new { left = Safe(() => setupApi.LeftFooter), center = Safe(() => setupApi.CenterFooter), right = Safe(() => setupApi.RightFooter) },
                horizontalPageBreaks = ReadPageBreaks(sheet, horizontal: true),
                verticalPageBreaks = ReadPageBreaks(sheet, horizontal: false),
            };
        }
        finally
        {
            ComInterop.Release(setup);
        }
    }

    private static List<string> ReadPageBreaks(dynamic sheet, bool horizontal)
    {
        var addresses = new List<string>();
        object? breaks = null;
        try
        {
            breaks = horizontal ? sheet.HPageBreaks : sheet.VPageBreaks;
            dynamic breaksApi = breaks;
            for (var index = 1; index <= Convert.ToInt32(breaksApi.Count); index++)
            {
                object? pageBreak = null;
                object? location = null;
                try
                {
                    pageBreak = breaksApi.Item(index);
                    location = ((dynamic)pageBreak).Location;
                    var address = Convert.ToString(((dynamic)location).Address[false, false]);
                    if (!string.IsNullOrWhiteSpace(address)) addresses.Add(address);
                }
                catch
                {
                    // Manual and automatic page breaks are not uniformly inspectable across hosts.
                }
                finally
                {
                    ComInterop.Release(location);
                    ComInterop.Release(pageBreak);
                }
            }
            return addresses;
        }
        finally
        {
            ComInterop.Release(breaks);
        }
    }

    private static bool TryNumber(JsonElement value, string name, out double result)
    {
        result = 0;
        return value.TryGetProperty(name, out var property) && property.ValueKind == JsonValueKind.Number && property.TryGetDouble(out result);
    }

    private static bool TryString(JsonElement value, string name, out string result)
    {
        result = string.Empty;
        if (!value.TryGetProperty(name, out var property) || property.ValueKind != JsonValueKind.String) return false;
        result = property.GetString() ?? string.Empty;
        return true;
    }

    private static object PrintSnapshot(dynamic setup) => new
    {
        area = Safe(() => setup.PrintArea),
        orientation = Convert.ToInt32(Safe(() => setup.Orientation) ?? 1) == 2 ? "landscape" : "portrait",
        paperSize = Safe(() => setup.PaperSize),
        fitToPagesWide = Safe(() => setup.FitToPagesWide),
        fitToPagesTall = Safe(() => setup.FitToPagesTall),
        repeatRows = Safe(() => setup.PrintTitleRows),
        repeatColumns = Safe(() => setup.PrintTitleColumns),
        header = Safe(() => setup.CenterHeader),
        footer = Safe(() => setup.CenterFooter),
    };

    private static int PaperSize(string value) => value.ToLowerInvariant() switch { "a3" => 8, "a4" => 9, "a5" => 11, "letter" => 1, "legal" => 5, _ => 9 };
    private static object? Safe(Func<object?> value) { try { return value(); } catch { return null; } }
}
