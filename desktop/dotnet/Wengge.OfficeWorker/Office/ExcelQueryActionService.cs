using System.Runtime.InteropServices;
using System.Text.Json;
using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Office;

internal sealed class ExcelQueryActionService(OfficeApplicationProvider applications)
{
    private static readonly HashSet<string> Operations = ["createPowerQuery", "inspectPowerQueries", "managePowerQuery"];

    public static bool Supports(string operation) => Operations.Contains(operation);

    public object Execute(OfficeActionRequest request)
    {
        using var context = new ExcelActionContext(applications, request);
        if (request.Operation == "inspectPowerQueries")
            return OfficeActionResults.Done(request, "com", "已检查 Power Query", Snapshot(context.Workbook, request.StringParam("name")), Array.Empty<OfficeChange>());
        var command = request.Operation == "createPowerQuery" ? "upsert" : request.StringParam("command", "upsert");
        var name = request.StringParam("name");
        if (name.Length == 0) throw new OfficeWorkerException("invalid_params", "Power Query 操作需要 params.name");
        object? query = FindQuery(context.Workbook, name);
        try
        {
            switch (command)
            {
                case "create":
                case "update":
                case "upsert":
                    query = Upsert(context.Workbook, query, name, request, command);
                    break;
                case "duplicate":
                    query = Duplicate(context.Workbook, query, request, ref name);
                    break;
                case "rename":
                    if (query is null) throw NotFound(name);
                    var newName = request.StringParam("newName");
                    if (newName.Length == 0) throw new OfficeWorkerException("invalid_params", "rename 需要 params.newName");
                    if (FindQuery(context.Workbook, newName) is { } existing) { ComInterop.Release(existing); throw new OfficeWorkerException("query_exists", $"Power Query 已存在: {newName}"); }
                    ((dynamic)query).Name = newName;
                    name = newName;
                    break;
                case "load":
                    if (query is null) throw NotFound(name);
                    ConfigureLoad(context, name, request);
                    break;
                case "refresh":
                    if (query is null) throw NotFound(name);
                    Refresh(context, name);
                    break;
                case "unload":
                    if (query is null) throw NotFound(name);
                    RemoveLoads(context.Workbook, name, request.BoolParam("clearOutput"));
                    break;
                case "delete":
                    if (query is null) throw NotFound(name);
                    RemoveLoads(context.Workbook, name, request.BoolParam("clearOutput"));
                    ((dynamic)query).Delete();
                    ComInterop.Release(query);
                    query = null;
                    break;
                default:
                    throw new OfficeWorkerException("unsupported_operation", $"不支持的 Power Query 命令: {command}");
            }
            if (request.StringParam("loadMode").Length > 0 && command is not ("load" or "refresh" or "unload" or "delete"))
                ConfigureLoad(context, name, request);
            if (request.BoolParam("refresh")) Refresh(context, name);
            context.Save(request);
            var data = new { command, queryName = name, snapshot = Snapshot(context.Workbook, command == "delete" ? string.Empty : name) };
            return OfficeActionResults.Done(request, "com", $"已执行 Power Query {command}", data,
                [new OfficeChange("power-query", name, $"已执行 Power Query {command}")]);
        }
        finally { ComInterop.Release(query); }
    }

    private static object Upsert(dynamic workbook, object? query, string name, OfficeActionRequest request, string command)
    {
        var formula = request.StringParam("mFormula");
        if (formula.Length == 0) throw new OfficeWorkerException("invalid_params", $"{command} 需要 params.mFormula");
        if (command == "create" && query is not null) throw new OfficeWorkerException("query_exists", $"Power Query 已存在: {name}");
        if (command == "update" && query is null) throw NotFound(name);
        if (query is null)
        {
            object? queries = null;
            try
            {
                queries = workbook.Queries;
                return ((dynamic)queries).Add(name, formula, request.StringParam("description"));
            }
            finally { ComInterop.Release(queries); }
        }
        dynamic queryApi = query;
        queryApi.Formula = formula;
        if (request.Param("description").ValueKind == JsonValueKind.String) queryApi.Description = request.StringParam("description");
        return query;
    }

    private static object Duplicate(dynamic workbook, object? query, OfficeActionRequest request, ref string name)
    {
        if (query is null) throw NotFound(name);
        var newName = request.StringParam("newName");
        if (newName.Length == 0) throw new OfficeWorkerException("invalid_params", "duplicate 需要 params.newName");
        object? existing = FindQuery(workbook, newName);
        if (existing is not null) { ComInterop.Release(existing); throw new OfficeWorkerException("query_exists", $"Power Query 已存在: {newName}"); }
        object? queries = null;
        try
        {
            queries = workbook.Queries;
            dynamic queryApi = query;
            name = newName;
            return ((dynamic)queries).Add(newName, Convert.ToString(queryApi.Formula), Convert.ToString(queryApi.Description));
        }
        finally { ComInterop.Release(queries); }
    }

    internal static object Snapshot(dynamic workbook, string onlyName)
    {
        Exception? lastError = null;
        for (var attempt = 1; attempt <= 3; attempt++)
        {
            try { return SnapshotOnce(workbook, onlyName); }
            catch (OfficeWorkerException exception) when (exception.Code == "power_query_unavailable") { throw; }
            catch (Exception exception)
            {
                lastError = exception;
                if (attempt < 3) Thread.Sleep(200 * attempt);
            }
        }
        throw new OfficeWorkerException(
            "power_query_inspection_failed",
            "读取 Power Query 详情失败，请稍后重试或检查查询连接状态",
            new { errorType = lastError?.GetType().FullName, hResult = lastError?.HResult, message = lastError?.Message },
            lastError);
    }

    private static object SnapshotOnce(dynamic workbook, string onlyName)
    {
        var queries = new List<object>();
        var connections = new List<object>();
        object? queryCollection = null;
        object? connectionCollection = null;
        try
        {
            try { queryCollection = workbook.Queries; }
            catch (Exception exception) when (MissingQueriesMember(exception))
            {
                throw new OfficeWorkerException(
                    "power_query_unavailable",
                    "当前宿主没有提供 Workbook.Queries COM 接口",
                    new { errorType = exception.GetType().FullName, hResult = exception.HResult },
                    exception);
            }
            dynamic queryCollectionApi = queryCollection;
            for (var index = 1; index <= Convert.ToInt32(queryCollectionApi.Count); index++)
            {
                object? query = null;
                try
                {
                    query = queryCollectionApi.Item(index);
                    dynamic queryApi = query;
                    var name = Convert.ToString(queryApi.Name) ?? string.Empty;
                    if (onlyName.Length > 0 && !name.Equals(onlyName, StringComparison.OrdinalIgnoreCase)) continue;
                    queries.Add(new
                    {
                        name,
                        formula = Convert.ToString(queryApi.Formula),
                        description = Safe(() => queryApi.Description),
                        loads = ReadLoads(workbook, name),
                    });
                }
                finally { ComInterop.Release(query); }
            }
            connectionCollection = workbook.Connections;
            dynamic connectionCollectionApi = connectionCollection;
            for (var index = 1; index <= Convert.ToInt32(connectionCollectionApi.Count); index++)
            {
                object? connection = null;
                try
                {
                    connection = connectionCollectionApi.Item(index);
                    dynamic connectionApi = connection;
                    var name = Convert.ToString(connectionApi.Name) ?? string.Empty;
                    if (onlyName.Length > 0 && !name.Equals($"Query - {onlyName}", StringComparison.OrdinalIgnoreCase)) continue;
                    connections.Add(new { name, type = Safe(() => connectionApi.Type), description = Safe(() => connectionApi.Description) });
                }
                finally { ComInterop.Release(connection); }
            }
            return new { queries, connections, queryCount = queries.Count };
        }
        finally
        {
            ComInterop.Release(connectionCollection);
            ComInterop.Release(queryCollection);
        }
    }

    private static bool MissingQueriesMember(Exception exception)
    {
        if (exception is MissingMemberException) return true;
        if (exception.GetType().FullName == "Microsoft.CSharp.RuntimeBinder.RuntimeBinderException") return true;
        if (exception is not COMException comException) return false;
        return comException.HResult is unchecked((int)0x80020003) or unchecked((int)0x80020006);
    }

    private static object? FindQuery(dynamic workbook, string name)
    {
        if (name.Length == 0) return null;
        object? queries = null;
        try
        {
            queries = workbook.Queries;
            try { return ((dynamic)queries).Item(name); } catch { return null; }
        }
        finally { ComInterop.Release(queries); }
    }

    private static List<object> ReadLoads(dynamic workbook, string name)
    {
        var loads = new List<object>();
        object? worksheets = null;
        object? connections = null;
        try
        {
            worksheets = workbook.Worksheets;
            dynamic worksheetsApi = worksheets;
            for (var sheetIndex = 1; sheetIndex <= Convert.ToInt32(worksheetsApi.Count); sheetIndex++)
            {
                object? sheet = null;
                object? tables = null;
                try
                {
                    sheet = worksheetsApi.Item(sheetIndex);
                    dynamic sheetApi = sheet;
                    tables = sheetApi.ListObjects;
                    dynamic tablesApi = tables;
                    for (var tableIndex = 1; tableIndex <= Convert.ToInt32(tablesApi.Count); tableIndex++)
                    {
                        object? table = null;
                        object? queryTable = null;
                        object? connection = null;
                        object? range = null;
                        try
                        {
                            table = tablesApi.Item(tableIndex);
                            dynamic tableApi = table;
                            queryTable = tableApi.QueryTable;
                            dynamic queryTableApi = queryTable;
                            try { connection = queryTableApi.WorkbookConnection; } catch { }
                            var connectionName = connection is null ? string.Empty : Convert.ToString(((dynamic)connection).Name) ?? string.Empty;
                            var commandText = SafeText(() => queryTableApi.CommandText);
                            if (!connectionName.Equals($"Query - {name}", StringComparison.OrdinalIgnoreCase)
                                && !commandText.Contains($"[{name}]", StringComparison.OrdinalIgnoreCase)) continue;
                            range = tableApi.Range;
                            loads.Add(new
                            {
                                kind = "worksheet",
                                sheet = Convert.ToString(sheetApi.Name),
                                table = Convert.ToString(tableApi.Name),
                                range = Safe(() => ((dynamic)range).Address[false, false]),
                                connection = connectionName,
                                refreshing = Safe(() => queryTableApi.Refreshing) ?? false,
                            });
                        }
                        catch
                        {
                            // Ordinary ListObjects do not expose QueryTable and are not Power Query loads.
                        }
                        finally
                        {
                            ComInterop.Release(range);
                            ComInterop.Release(connection);
                            ComInterop.Release(queryTable);
                            ComInterop.Release(table);
                        }
                    }
                }
                finally
                {
                    ComInterop.Release(tables);
                    ComInterop.Release(sheet);
                }
            }

            connections = workbook.Connections;
            dynamic connectionsApi = connections;
            for (var index = 1; index <= Convert.ToInt32(connectionsApi.Count); index++)
            {
                object? connection = null;
                object? modelConnection = null;
                try
                {
                    connection = connectionsApi.Item(index);
                    dynamic connectionApi = connection;
                    var connectionName = Convert.ToString(connectionApi.Name) ?? string.Empty;
                    if (!connectionName.Equals($"Query - {name}", StringComparison.OrdinalIgnoreCase)) continue;
                    try { modelConnection = connectionApi.ModelConnection; } catch { }
                    loads.Add(new { kind = modelConnection is null ? "connection" : "dataModel", connection = connectionName });
                }
                finally
                {
                    ComInterop.Release(modelConnection);
                    ComInterop.Release(connection);
                }
            }
            return loads;
        }
        finally
        {
            ComInterop.Release(connections);
            ComInterop.Release(worksheets);
        }
    }

    private static string SafeText(Func<object?> value)
    {
        try
        {
            var result = value();
            if (result is not Array values) return Convert.ToString(result) ?? string.Empty;
            return string.Concat(values.Cast<object?>().Select(item => Convert.ToString(item)));
        }
        catch
        {
            return string.Empty;
        }
    }

    private static void ConfigureLoad(ExcelActionContext context, string name, OfficeActionRequest request)
    {
        RemoveLoads(context.Workbook, name, request.BoolParam("clearOutput"));
        var mode = request.StringParam("loadMode", "worksheet");
        if (mode == "worksheet") AddWorksheetLoad(context, name, request);
        else AddConnection(context.Workbook, name, mode == "dataModel");
    }

    private static void AddConnection(dynamic workbook, string name, bool dataModel)
    {
        object? connections = null;
        try
        {
            connections = workbook.Connections;
            dynamic connectionsApi = connections;
            try { ((dynamic)connectionsApi.Item($"Query - {name}")).Delete(); } catch { }
            var connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={name};Extended Properties=\"\"";
            connectionsApi.Add2($"Query - {name}", $"Power Query - {name}", connectionString, $"SELECT * FROM [{name.Replace("]", "]]", StringComparison.Ordinal)}]", 2, dataModel, false);
        }
        finally { ComInterop.Release(connections); }
    }

    private static void AddWorksheetLoad(ExcelActionContext context, string name, OfficeActionRequest request)
    {
        var destination = request.StringParam("destination");
        if (destination.Length == 0) throw new OfficeWorkerException("invalid_params", "加载到工作表需要 params.destination，例如 QueryOutput!A1");
        var parts = destination.Split('!', 2);
        var (defaultSheet, _) = request.ExcelTarget();
        object? sheet = null;
        object? start = null;
        object? listObjects = null;
        object? table = null;
        object? queryTable = null;
        try
        {
            sheet = context.Workbook.Worksheets.Item(parts.Length == 2 ? parts[0].Trim('\'') : defaultSheet);
            start = ((dynamic)sheet).Range(parts.Length == 2 ? parts[1] : destination);
            listObjects = ((dynamic)sheet).ListObjects;
            var connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={name};Extended Properties=\"\"";
            table = ((dynamic)listObjects).Add(0, new object[] { connectionString }, Type.Missing, 1, start);
            var tableName = request.StringParam("tableName");
            if (tableName.Length > 0) ((dynamic)table).Name = tableName;
            queryTable = ((dynamic)table).QueryTable;
            dynamic queryTableApi = queryTable;
            queryTableApi.CommandType = 2;
            queryTableApi.CommandText = new object[] { $"SELECT * FROM [{name.Replace("]", "]]", StringComparison.Ordinal)}]" };
            queryTableApi.Refresh(false);
        }
        finally
        {
            ComInterop.Release(queryTable);
            ComInterop.Release(table);
            ComInterop.Release(listObjects);
            ComInterop.Release(start);
            ComInterop.Release(sheet);
        }
    }

    private static int RemoveLoads(dynamic workbook, string name, bool clearOutput)
    {
        var removed = 0;
        object? worksheets = null;
        object? connections = null;
        try
        {
            worksheets = workbook.Worksheets;
            dynamic worksheetsApi = worksheets;
            for (var sheetIndex = 1; sheetIndex <= Convert.ToInt32(worksheetsApi.Count); sheetIndex++)
            {
                object? sheet = null;
                object? tables = null;
                try
                {
                    sheet = worksheetsApi.Item(sheetIndex);
                    tables = ((dynamic)sheet).ListObjects;
                    dynamic tablesApi = tables;
                    for (var index = Convert.ToInt32(tablesApi.Count); index >= 1; index--)
                    {
                        object? table = null;
                        object? queryTable = null;
                        object? connection = null;
                        object? range = null;
                        try
                        {
                            table = tablesApi.Item(index);
                            queryTable = ((dynamic)table).QueryTable;
                            connection = ((dynamic)queryTable).WorkbookConnection;
                            if (!string.Equals(Convert.ToString(((dynamic)connection).Name), $"Query - {name}", StringComparison.OrdinalIgnoreCase)) continue;
                            range = ((dynamic)table).Range;
                            ((dynamic)table).Unlist();
                            if (clearOutput) ((dynamic)range).Clear();
                            removed++;
                        }
                        catch { }
                        finally { ComInterop.Release(range); ComInterop.Release(connection); ComInterop.Release(queryTable); ComInterop.Release(table); }
                    }
                }
                finally { ComInterop.Release(tables); ComInterop.Release(sheet); }
            }
            connections = workbook.Connections;
            dynamic connectionsApi = connections;
            for (var index = Convert.ToInt32(connectionsApi.Count); index >= 1; index--)
            {
                object? connection = null;
                try
                {
                    connection = connectionsApi.Item(index);
                    if (!string.Equals(Convert.ToString(((dynamic)connection).Name), $"Query - {name}", StringComparison.OrdinalIgnoreCase)) continue;
                    ((dynamic)connection).Delete();
                    removed++;
                }
                finally { ComInterop.Release(connection); }
            }
            return removed;
        }
        finally { ComInterop.Release(connections); ComInterop.Release(worksheets); }
    }

    private static void Refresh(ExcelActionContext context, string name)
    {
        object? connections = null;
        try
        {
            connections = context.Workbook.Connections;
            dynamic connectionsApi = connections;
            for (var index = 1; index <= Convert.ToInt32(connectionsApi.Count); index++)
            {
                object? connection = null;
                try
                {
                    connection = connectionsApi.Item(index);
                    if (string.Equals(Convert.ToString(((dynamic)connection).Name), $"Query - {name}", StringComparison.OrdinalIgnoreCase)) ((dynamic)connection).Refresh();
                }
                finally { ComInterop.Release(connection); }
            }
            try { context.App.CalculateUntilAsyncQueriesDone(); } catch { }
        }
        finally { ComInterop.Release(connections); }
    }

    private static OfficeWorkerException NotFound(string name) => new("query_not_found", $"找不到 Power Query: {name}");
    private static object? Safe(Func<object?> value) { try { return value(); } catch { return null; } }
}
