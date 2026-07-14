using System.Text.Json;
using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Office;

internal sealed class ExcelObjectActionService(OfficeApplicationProvider applications)
{
    public static bool Supports(string operation) => operation is "inspectWorkbookObjects" or "manageWorkbookObject" or "manageWorksheetObjects";

    public object Execute(OfficeActionRequest request)
    {
        using var context = new ExcelActionContext(applications, request);
        if (request.Operation == "inspectWorkbookObjects")
            return OfficeActionResults.Done(request, "com", "已检查工作簿对象", new { objects = Inspect(context.Workbook) }, Array.Empty<OfficeChange>());
        var objectType = request.Operation == "manageWorksheetObjects"
            ? "shape"
            : request.StringParam("objectType", request.StringParam("type", "name"));
        var command = request.StringParam("command", "update");
        _ = ManageObject(context, request, objectType);
        context.Save(request);
        var data = new { objectType, command, objects = Inspect(context.Workbook) };
        return OfficeActionResults.Done(request, "com", "已更新工作簿对象", data,
            [new OfficeChange("workbook-object", request.StringParam("name", request.ExcelTarget().SheetName), "已更新工作簿对象")]);
    }

    private static object Inspect(dynamic workbook)
    {
        var worksheets = new List<object>();
        var names = new List<object>();
        var tables = new List<object>();
        var charts = new List<object>();
        var shapes = new List<object>();
        object? sheets = null;
        object? workbookNames = null;
        try
        {
            sheets = workbook.Worksheets;
            dynamic sheetsApi = sheets;
            for (var index = 1; index <= Convert.ToInt32(sheetsApi.Count); index++)
            {
                object? sheet = null;
                object? usedRange = null;
                object? listObjects = null;
                object? chartObjects = null;
                object? shapeCollection = null;
                try
                {
                    sheet = sheetsApi.Item(index);
                    dynamic sheetApi = sheet;
                    usedRange = sheetApi.UsedRange;
                    listObjects = sheetApi.ListObjects;
                    chartObjects = sheetApi.ChartObjects();
                    shapeCollection = sheetApi.Shapes;
                    worksheets.Add(new
                    {
                        name = Convert.ToString(sheetApi.Name), index,
                        visible = Convert.ToInt32(sheetApi.Visible),
                        usedRange = Safe(() => ((dynamic)usedRange).Address[false, false]),
                        tableCount = Convert.ToInt32(((dynamic)listObjects).Count),
                        chartCount = Convert.ToInt32(((dynamic)chartObjects).Count),
                        shapeCount = Convert.ToInt32(((dynamic)shapeCollection).Count),
                    });
                    ReadTables(listObjects, sheetApi, tables);
                    ReadCharts(chartObjects, sheetApi, charts);
                    ReadShapes(shapeCollection, sheetApi, shapes);
                }
                finally
                {
                    ComInterop.Release(shapeCollection);
                    ComInterop.Release(chartObjects);
                    ComInterop.Release(listObjects);
                    ComInterop.Release(usedRange);
                    ComInterop.Release(sheet);
                }
            }
            workbookNames = workbook.Names;
            dynamic namesApi = workbookNames;
            for (var index = 1; index <= Convert.ToInt32(namesApi.Count); index++)
            {
                object? name = null;
                try
                {
                    name = namesApi.Item(index);
                    dynamic nameApi = name;
                    names.Add(new { name = Convert.ToString(nameApi.Name), refersTo = Safe(() => nameApi.RefersTo), visible = Safe(() => nameApi.Visible), scope = "workbook" });
                }
                finally { ComInterop.Release(name); }
            }
            return new { worksheets, names, tables, charts, shapes };
        }
        finally { ComInterop.Release(workbookNames); ComInterop.Release(sheets); }
    }

    private static object ManageObject(ExcelActionContext context, OfficeActionRequest request, string? objectTypeOverride = null)
    {
        var objectType = objectTypeOverride ?? request.StringParam("objectType", request.StringParam("type", "name"));
        var command = request.StringParam("command", "update");
        return objectType switch
        {
            "worksheet" => ManageWorksheet(context, request),
            "name" => ManageName(context.Workbook, request, command),
            "table" => ManageTable(context, request, command),
            "chart" => ManageChartOrShape(context, request, command, chart: true),
            "shape" => ManageChartOrShape(context, request, command, chart: false),
            "connection" => ManageConnection(context.Workbook, request, command),
            _ => throw new OfficeWorkerException("unsupported_operation", $"不支持的工作簿对象类型: {objectType}"),
        };
    }

    private static object ManageWorksheet(ExcelActionContext context, OfficeActionRequest request)
    {
        var command = request.StringParam("command", "update");
        var name = request.StringParam("sheetName", request.StringParam("name", request.ExcelTarget().SheetName));
        object? sheets = null;
        object? sheet = null;
        try
        {
            sheets = context.Workbook.Worksheets;
            dynamic sheetsApi = sheets;
            if (command == "add")
            {
                sheet = sheetsApi.Add();
                ((dynamic)sheet).Name = name;
                return new { command, name };
            }
            sheet = sheetsApi.Item(name);
            dynamic sheetApi = sheet;
            switch (command)
            {
                case "delete":
                    var alerts = context.App.DisplayAlerts;
                    try { context.App.DisplayAlerts = false; sheetApi.Delete(); }
                    finally { context.App.DisplayAlerts = alerts; }
                    break;
                case "rename": sheetApi.Name = request.StringParam("newName"); break;
                case "copy": sheetApi.Copy(After: sheetsApi.Item(sheetsApi.Count)); break;
                case "move": sheetApi.Move(Before: sheetsApi.Item(Math.Max(1, request.IntParam("position", 1)))); break;
                case "hide": sheetApi.Visible = 0; break;
                case "veryHide": sheetApi.Visible = 2; break;
                case "show": sheetApi.Visible = -1; break;
                case "protect": sheetApi.Protect(request.StringParam("password")); break;
                case "unprotect": sheetApi.Unprotect(request.StringParam("password")); break;
                case "update":
                    if (request.Param("visible").ValueKind is JsonValueKind.True or JsonValueKind.False) sheetApi.Visible = request.BoolParam("visible") ? -1 : 0;
                    if (request.StringParam("tabColor").Length > 0) sheetApi.Tab.Color = ExcelActionService.OleColor(request.StringParam("tabColor"));
                    break;
                default: throw new OfficeWorkerException("unsupported_operation", $"不支持的工作表命令: {command}");
            }
            return new { command, name };
        }
        finally { ComInterop.Release(sheet); ComInterop.Release(sheets); }
    }

    private static object ManageName(dynamic workbook, OfficeActionRequest request, string command)
    {
        var name = request.StringParam("name");
        if (name.Length == 0) throw new OfficeWorkerException("invalid_params", "名称对象操作需要 params.name");
        object? names = null;
        object? definedName = null;
        try
        {
            names = workbook.Names;
            dynamic namesApi = names;
            try { definedName = namesApi.Item(name); } catch { }
            if (command == "delete")
            {
                if (definedName is null) throw new OfficeWorkerException("object_not_found", $"找不到名称: {name}");
                ((dynamic)definedName).Delete();
            }
            else if (definedName is null)
            {
                var refersTo = request.StringParam("refersTo");
                if (refersTo.Length == 0) throw new OfficeWorkerException("invalid_params", "创建名称需要 params.refersTo");
                definedName = namesApi.Add(name, refersTo);
            }
            else
            {
                dynamic nameApi = definedName;
                if (request.StringParam("refersTo").Length > 0) nameApi.RefersTo = request.StringParam("refersTo");
                if (request.StringParam("newName").Length > 0) nameApi.Name = request.StringParam("newName");
                if (request.Param("visible").ValueKind is JsonValueKind.True or JsonValueKind.False) nameApi.Visible = request.BoolParam("visible");
            }
            return new { command, name };
        }
        finally { ComInterop.Release(definedName); ComInterop.Release(names); }
    }

    private static object ManageTable(ExcelActionContext context, OfficeActionRequest request, string command)
    {
        var (sheet, range) = context.GetRange(request);
        object? tables = null;
        object? table = null;
        try
        {
            tables = ((dynamic)sheet).ListObjects;
            dynamic tablesApi = tables;
            var name = request.StringParam("name");
            if (name.Length > 0) { try { table = tablesApi.Item(name); } catch { } }
            if (command == "delete")
            {
                if (table is null) throw new OfficeWorkerException("object_not_found", $"找不到表格: {name}");
                if (request.BoolParam("unlist")) ((dynamic)table).Unlist(); else ((dynamic)table).Delete();
            }
            else if (table is null)
            {
                table = tablesApi.Add(1, range, Type.Missing, 1, Type.Missing);
                if (name.Length > 0) ((dynamic)table).Name = name;
            }
            if (table is not null && command != "delete")
            {
                dynamic tableApi = table;
                if (request.StringParam("newName").Length > 0) tableApi.Name = request.StringParam("newName");
                if (request.StringParam("style").Length > 0) tableApi.TableStyle = request.StringParam("style");
                if (request.Param("showTotals").ValueKind is JsonValueKind.True or JsonValueKind.False) tableApi.ShowTotals = request.BoolParam("showTotals");
            }
            return new { command, name };
        }
        finally { ComInterop.Release(table); ComInterop.Release(tables); ComInterop.Release(range); ComInterop.Release(sheet); }
    }

    private static object ManageChartOrShape(ExcelActionContext context, OfficeActionRequest request, string command, bool chart)
    {
        var (sheet, range) = context.GetRange(request);
        object? collection = null;
        object? item = null;
        try
        {
            collection = chart ? ((dynamic)sheet).ChartObjects() : ((dynamic)sheet).Shapes;
            dynamic collectionApi = collection;
            var name = request.StringParam("name");
            item = name.Length > 0 ? collectionApi.Item(name) : collectionApi.Item(Math.Max(1, request.IntParam("index", 1)));
            dynamic itemApi = item;
            if (command == "delete") itemApi.Delete();
            else
            {
                if (request.StringParam("newName").Length > 0) itemApi.Name = request.StringParam("newName");
                if (request.Param("left").ValueKind == JsonValueKind.Number) itemApi.Left = request.DoubleParam("left");
                if (request.Param("top").ValueKind == JsonValueKind.Number) itemApi.Top = request.DoubleParam("top");
                if (request.Param("width").ValueKind == JsonValueKind.Number) itemApi.Width = request.DoubleParam("width");
                if (request.Param("height").ValueKind == JsonValueKind.Number) itemApi.Height = request.DoubleParam("height");
            }
            return new { command, name, objectType = chart ? "chart" : "shape" };
        }
        finally { ComInterop.Release(item); ComInterop.Release(collection); ComInterop.Release(range); ComInterop.Release(sheet); }
    }

    private static object ManageConnection(dynamic workbook, OfficeActionRequest request, string command)
    {
        var name = request.StringParam("name");
        object? connections = null;
        object? connection = null;
        try
        {
            connections = workbook.Connections;
            connection = ((dynamic)connections).Item(name);
            if (command == "delete") ((dynamic)connection).Delete();
            else if (command == "refresh") ((dynamic)connection).Refresh();
            else throw new OfficeWorkerException("unsupported_operation", $"不支持的连接命令: {command}");
            return new { command, name };
        }
        finally { ComInterop.Release(connection); ComInterop.Release(connections); }
    }

    private static void ReadTables(dynamic tables, dynamic sheet, List<object> output)
    {
        for (var index = 1; index <= Convert.ToInt32(tables.Count); index++)
        {
            object? table = null;
            object? range = null;
            try
            {
                table = tables.Item(index); dynamic item = table; range = item.Range;
                output.Add(new { sheet = Convert.ToString(sheet.Name), name = Convert.ToString(item.Name), displayName = Convert.ToString(item.DisplayName), range = Safe(() => ((dynamic)range).Address[false, false]), style = Safe(() => item.TableStyle), showTotals = Safe(() => item.ShowTotals) });
            }
            finally { ComInterop.Release(range); ComInterop.Release(table); }
        }
    }

    private static void ReadCharts(dynamic charts, dynamic sheet, List<object> output)
    {
        for (var index = 1; index <= Convert.ToInt32(charts.Count); index++)
        {
            object? chart = null;
            try { chart = charts.Item(index); dynamic item = chart; output.Add(new { sheet = Convert.ToString(sheet.Name), name = Convert.ToString(item.Name), left = item.Left, top = item.Top, width = item.Width, height = item.Height }); }
            finally { ComInterop.Release(chart); }
        }
    }

    private static void ReadShapes(dynamic shapes, dynamic sheet, List<object> output)
    {
        for (var index = 1; index <= Convert.ToInt32(shapes.Count); index++)
        {
            object? shape = null;
            try { shape = shapes.Item(index); dynamic item = shape; output.Add(new { sheet = Convert.ToString(sheet.Name), name = Convert.ToString(item.Name), type = Safe(() => item.Type), left = item.Left, top = item.Top, width = item.Width, height = item.Height }); }
            finally { ComInterop.Release(shape); }
        }
    }

    private static object? Safe(Func<object?> value) { try { return value(); } catch { return null; } }
}
