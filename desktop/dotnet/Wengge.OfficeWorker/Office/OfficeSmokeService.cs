using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Excel;
using Wengge.OfficeWorker.Protocol;
using Wengge.OfficeWorker.Runtime;

namespace Wengge.OfficeWorker.Office;

internal sealed class OfficeSmokeService(OfficeDocumentService documents, ExcelSessionService excelSessions) : IDisposable
{
    private readonly Dictionary<int, List<object>> ownedApplications = [];

    public object GetExcelDisplayAlerts()
    {
        EnsureEnabled();
        using var handle = excelSessions.GetActiveRequired();
        dynamic app = handle.Application;
        // Quarantine COM dynamic at the acquisition boundary before ToBoolean.
        object? displayAlerts = app.DisplayAlerts;
        return new { displayAlerts = ToBoolean(displayAlerts) };
    }

    public object SetExcelDisplayAlerts(JsonElement parameters)
    {
        EnsureEnabled();
        var value = RequiredBoolean(parameters, "displayAlerts");
        using var handle = excelSessions.GetActiveRequired();
        dynamic app = handle.Application;
        app.DisplayAlerts = value;
        object? displayAlerts = app.DisplayAlerts;
        return new { displayAlerts = ToBoolean(displayAlerts) };
    }

    public object SetExcelStructureProtected(JsonElement parameters)
    {
        EnsureEnabled();
        var protect = RequiredBoolean(parameters, "protected");
        var password = parameters.OptionalString("password") ?? "wengge-smoke-m09";
        using var handle = excelSessions.GetActiveRequired();
        dynamic app = handle.Application;
        object? workbook = null;
        try
        {
            workbook = app.ActiveWorkbook ?? throw new OfficeWorkerException("workbook_not_found", "当前没有活动工作簿");
            dynamic workbookApi = workbook;
            if (protect)
            {
                // Workbook.Protect(Password, Structure, Windows)
                workbookApi.Protect(password, true, false);
            }
            else
            {
                workbookApi.Unprotect(password);
            }

            object? structureProtected = workbookApi.ProtectStructure;
            return new
            {
                structureProtected = ToBoolean(structureProtected),
                passwordUsed = !string.IsNullOrEmpty(password),
            };
        }
        finally
        {
            ComInterop.Release(workbook);
        }
    }

    public object MarkWordBookmarkDirty(JsonElement parameters)
    {
        EnsureEnabled();
        var filePath = parameters.RequiredString("filePath");
        var instanceId = parameters.RequiredString("instanceId");
        var name = parameters.RequiredString("name");
        var text = parameters.RequiredString("text");
        return documents.WithDocument("word", filePath, instanceId, null, 0, handle =>
        {
            dynamic document = handle.Document;
            var start = Math.Max(0, OfficeDocumentService.SafeInt32(() => document.Content.End) - 1);
            object? range = null;
            object? bookmarkRange = null;
            try
            {
                range = document.Range(start, start);
                ((dynamic)range).InsertAfter(text);
                if (OfficeDocumentService.SafeBoolean(() => document.Bookmarks.Exists(name), false)) document.Bookmarks.Item(name).Delete();
                bookmarkRange = document.Range(start, start + text.Length);
                document.Bookmarks.Add(name, bookmarkRange);
                return new { saved = OfficeDocumentService.SafeBoolean(() => document.Saved, true), instanceId = handle.InstanceId };
            }
            finally { ComInterop.Release(bookmarkRange); ComInterop.Release(range); }
        });
    }

    public object OpenFixtures(JsonElement parameters)
    {
        EnsureEnabled();
        var before = MicrosoftProcessIds().ToHashSet();
        var opened = new Dictionary<string, List<int>>
        {
            ["excel"] = [], ["word"] = [], ["presentation"] = [],
        };
        try
        {
            OpenEach(parameters.PropertyOrEmpty("excelPaths"), "excel", "Excel.Application", opened["excel"]);
            OpenEach(parameters.PropertyOrEmpty("wordPaths"), "word", "Word.Application", opened["word"]);
            OpenEach(parameters.PropertyOrEmpty("presentationPaths"), "presentation", "PowerPoint.Application", opened["presentation"]);
            foreach (var key in opened.Keys.ToArray()) opened[key] = opened[key].Where(id => id > 0 && !before.Contains(id)).Distinct().ToList();
            return new { excel = opened["excel"], word = opened["word"], presentation = opened["presentation"] };
        }
        catch
        {
            CloseOwned(ownedApplications.Keys.ToArray());
            throw;
        }
    }

    public object CloseFixtures(JsonElement parameters)
    {
        EnsureEnabled();
        var ids = new HashSet<int>();
        foreach (var name in new[] { "excel", "word", "presentation" })
            foreach (var id in ReadIntegers(parameters.PropertyOrEmpty(name))) if (id > 0) ids.Add(id);
        CloseOwned(ids);
        return new { closed = ids.Order().ToArray() };
    }

    public object ListProcesses()
    {
        EnsureEnabled();
        return new
        {
            microsoft = MicrosoftProcessIds(),
            wpsVisible = ProcessIds(["wps", "et", "wpp"], visibleOnly: true),
            // Spreadsheet hosts only (et/wps, including hidden); exclude presentation wpp.
            wpsAll = ProcessIds(["wps", "et"], visibleOnly: false),
        };
    }

    public object RunningProcesses(JsonElement ids)
    {
        EnsureEnabled();
        var running = Process.GetProcesses().Select(process => process.Id).ToHashSet();
        return ReadIntegers(ids).Where(running.Contains).Distinct().Order().ToArray();
    }

    public void Dispose()
    {
        if (ownedApplications.Count == 0) return;
        CloseOwned(ownedApplications.Keys.ToArray());
    }

    private void OpenEach(JsonElement paths, string app, string progId, List<int> processIds)
    {
        foreach (var path in ReadStrings(paths))
        {
            if (!File.Exists(Path.GetFullPath(path))) throw new OfficeWorkerException("file_not_found", $"冒烟文件不存在: {path}");
            var application = CreateApplication(progId);
            try
            {
                dynamic api = application;
                if (app == "presentation") api.Visible = -1;
                else api.Visible = true;
                if (app == "word") api.DisplayAlerts = 0;
                object? document = app switch
                {
                    "excel" => api.Workbooks.Open(Path.GetFullPath(path)),
                    "word" => api.Documents.Open(Path.GetFullPath(path)),
                    _ => api.Presentations.Open(Path.GetFullPath(path)),
                };
                ComInterop.Release(document);
                var processId = ProcessId(application);
                if (!ownedApplications.TryGetValue(processId, out var applications)) ownedApplications[processId] = applications = [];
                applications.Add(application);
                processIds.Add(processId);
            }
            catch
            {
                try { ((dynamic)application).Quit(); } catch { }
                ComInterop.Release(application);
                throw;
            }
        }
    }

    private static object CreateApplication(string progId)
    {
        Exception? lastError = null;
        for (var attempt = 1; attempt <= 3; attempt++)
        {
            object? application = null;
            try
            {
                application = ComInterop.CreateObject(progId);
                _ = Convert.ToString(((dynamic)application).Version);
                return application;
            }
            catch (Exception exception)
            {
                lastError = exception;
                if (application is not null) { try { ((dynamic)application).Quit(); } catch { } ComInterop.Release(application); }
                Thread.Sleep(300 * attempt);
            }
        }
        throw new OfficeWorkerException("office_unavailable", $"无法创建冒烟 COM 应用: {progId}", null, lastError);
    }

    private void CloseOwned(IEnumerable<int> processIds)
    {
        foreach (var processId in processIds.Distinct().ToArray())
        {
            if (ownedApplications.Remove(processId, out var applications))
            {
                foreach (var application in applications)
                {
                    try { ((dynamic)application).Quit(); } catch { }
                    ComInterop.Release(application);
                }
            }
            try
            {
                var process = Process.GetProcessById(processId);
                if (!process.WaitForExit(2_000)) process.Kill(entireProcessTree: true);
            }
            catch { }
        }
    }

    private static int ProcessId(object application)
    {
        var hwnd = OfficeDocumentService.SafeInt64(() => ((dynamic)application).Hwnd);
        if (hwnd == 0) hwnd = OfficeDocumentService.SafeInt64(() => ((dynamic)application).HWND);
        if (hwnd == 0) hwnd = OfficeDocumentService.SafeInt64(() => ((dynamic)application).ActiveWindow.Hwnd);
        if (hwnd == 0) hwnd = OfficeDocumentService.SafeInt64(() => ((dynamic)application).ActiveWindow.HWND);
        if (hwnd == 0) return 0;
        _ = GetWindowThreadProcessId(new IntPtr(hwnd), out var processId);
        return unchecked((int)processId);
    }

    private static int[] MicrosoftProcessIds() => ProcessIds(["EXCEL", "WINWORD", "POWERPNT"], visibleOnly: false);

    private static int[] ProcessIds(IEnumerable<string> names, bool visibleOnly)
    {
        var result = new HashSet<int>();
        foreach (var name in names)
            foreach (var process in Process.GetProcessesByName(name))
                try { if (!visibleOnly || process.MainWindowHandle != IntPtr.Zero) result.Add(process.Id); } finally { process.Dispose(); }
        return result.Order().ToArray();
    }

    private static IReadOnlyList<string> ReadStrings(JsonElement values) => values.ValueKind == JsonValueKind.Array
        ? values.EnumerateArray().Where(value => value.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(value.GetString())).Select(value => value.GetString()!).ToArray()
        : [];

    private static IReadOnlyList<int> ReadIntegers(JsonElement values) => values.ValueKind == JsonValueKind.Array
        ? values.EnumerateArray().Select(value => value.TryGetInt32(out var number) ? number : 0).Where(number => number > 0).ToArray()
        : [];

    private static void EnsureEnabled()
    {
        if (Environment.GetEnvironmentVariable("WENGGE_OFFICE_SMOKE") != "1")
            throw new OfficeWorkerException("smoke_disabled", "Office smoke RPC 仅在 WENGGE_OFFICE_SMOKE=1 时可用");
    }

    private static bool RequiredBoolean(JsonElement parameters, string name)
    {
        if (parameters.ValueKind != JsonValueKind.Object || !parameters.TryGetProperty(name, out var value)
            || value.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
        {
            throw new OfficeWorkerException("invalid_params", $"缺少参数: {name}");
        }

        return value.GetBoolean();
    }

    private static bool ToBoolean(object? value) => value switch
    {
        bool boolean => boolean,
        sbyte or byte or short or ushort or int or uint or long or ulong => Convert.ToInt64(value) != 0,
        _ => Convert.ToBoolean(value),
    };

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
}
