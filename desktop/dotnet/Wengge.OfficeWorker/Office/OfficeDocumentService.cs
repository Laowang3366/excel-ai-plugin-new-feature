using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.Json;
using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Protocol;
using Wengge.OfficeWorker.Runtime;

namespace Wengge.OfficeWorker.Office;

internal sealed class OfficeDocumentService
{
    private static readonly OfficeApplicationConfig[] ApplicationConfigs =
    [
        new("excel", ["Excel.Application", "Ket.Application"], "Workbooks"),
        new("word", ["Word.Application", "Kwps.Application", "Wps.Application"], "Documents"),
        new("presentation", ["PowerPoint.Application", "Wpp.Application", "Kwpp.Application"], "Presentations"),
    ];

    public object ListDocuments(string? appFilter = null)
    {
        ValidateApp(appFilter, allowEmpty: true);
        var handles = EnumerateDocuments(appFilter);
        try { return handles.Select(handle => Metadata(handle)).ToArray(); }
        finally { ReleaseHandles(handles); }
    }

    public object ActivateDocument(JsonElement parameters)
    {
        var app = parameters.RequiredString("app");
        ValidateApp(app);
        return WithDocument(
            app,
            parameters.OptionalString("filePath"),
            parameters.OptionalString("instanceId"),
            parameters.OptionalString("name"),
            parameters.OptionalInt32("index"),
            handle =>
            {
                dynamic application = handle.Application;
                dynamic document = handle.Document;
                try
                {
                    if (app == "presentation") application.Visible = -1;
                    else application.Visible = true;
                }
                catch { }
                document.Activate();
                return Metadata(handle, activeOverride: true);
            });
    }

    public object PrepareTransaction(JsonElement filePaths)
    {
        var wanted = ReadStringArray(filePaths).Select(PathKey).ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (wanted.Count == 0) return Array.Empty<object>();
        var handles = EnumerateDocuments(null);
        try
        {
            var prepared = new List<object>();
            var failures = new List<string>();
            foreach (var handle in handles)
            {
                var fullName = SafeString(() => ((dynamic)handle.Document).FullName);
                if (string.IsNullOrWhiteSpace(fullName) || !wanted.Contains(PathKey(fullName))) continue;
                var wasDirty = !SafeBoolean(() => ((dynamic)handle.Document).Saved, true);
                var saved = true;
                if (wasDirty)
                {
                    try { ((dynamic)handle.Document).Save(); }
                    catch { saved = false; failures.Add(fullName); }
                }
                prepared.Add(new
                {
                    app = handle.App,
                    filePath = fullName,
                    instanceId = handle.InstanceId,
                    wasDirty,
                    saved,
                });
            }
            if (failures.Count > 0)
                throw new OfficeWorkerException("transaction_prepare_failed", $"无法保存 {failures.Count} 个已打开的 Office 文档，事务已停止: {string.Join(", ", failures)}");
            return prepared.ToArray();
        }
        finally { ReleaseHandles(handles); }
    }

    public object RestoreTransactionFiles(JsonElement filesElement)
    {
        var files = ReadRestoreFiles(filesElement);
        if (files.Count == 0) return Array.Empty<object>();
        var restorePaths = files.Select(file => PathKey(file.FilePath)).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var handles = EnumerateDocuments(null);
        var sessions = new List<OpenDocumentSession>();
        try
        {
            foreach (var handle in handles)
            {
                var fullName = SafeString(() => ((dynamic)handle.Document).FullName);
                if (string.IsNullOrWhiteSpace(fullName) || !restorePaths.Contains(PathKey(fullName))) continue;
                dynamic document = handle.Document;
                if (!SafeBoolean(() => document.Saved, true)) document.Save();
                sessions.Add(new OpenDocumentSession(
                    handle.App,
                    fullName,
                    handle.InstanceId,
                    handle.Application,
                    IsActive(handle),
                    SafeBoolean(() => document.ReadOnly, false)));
                document.Close(false);
            }

            PublishRestoreFiles(files);

            var restored = new List<object>();
            var failures = new List<string>();
            foreach (var session in sessions)
            {
                var reopened = false;
                object? document = null;
                try
                {
                    if (File.Exists(session.FilePath))
                    {
                        dynamic app = session.Application;
                        document = session.App switch
                        {
                            "excel" => app.Workbooks.Open(session.FilePath, 0, session.ReadOnly),
                            "word" => app.Documents.Open(session.FilePath, false, session.ReadOnly),
                            _ => app.Presentations.Open(session.FilePath, session.ReadOnly, false, -1),
                        };
                        reopened = document is not null;
                        if (reopened && session.Active) ((dynamic)document!).Activate();
                    }
                }
                catch { reopened = false; }
                finally { ComInterop.Release(document); }
                restored.Add(new { app = session.App, filePath = session.FilePath, instanceId = session.InstanceId, reopened });
                if (!reopened && files.Any(file => file.Existed && PathKey(file.FilePath) == PathKey(session.FilePath)))
                    failures.Add(session.FilePath);
            }
            if (failures.Count > 0)
                throw new OfficeWorkerException("transaction_reopen_failed", $"事务文件已恢复，但 {failures.Count} 个 Office 文档无法重新打开: {string.Join(", ", failures)}");
            return restored.ToArray();
        }
        finally { ReleaseHandles(handles); }
    }

    internal T WithDocument<T>(
        string app,
        string? filePath,
        string? instanceId,
        string? name,
        int index,
        Func<OfficeDocumentHandle, T> operation)
    {
        using var lease = AcquireDocument(app, filePath, instanceId, name, index);
        return operation(lease.Handle);
    }

    internal static OfficeDocumentLease AcquireDocument(
        string app,
        string? filePath,
        string? instanceId,
        string? name = null,
        int index = 0)
    {
        ValidateApp(app);
        var handles = EnumerateDocuments(app);
        try
        {
            IEnumerable<OfficeDocumentHandle> candidates = handles;
            if (!string.IsNullOrWhiteSpace(instanceId))
                candidates = candidates.Where(handle => string.Equals(handle.InstanceId, instanceId, StringComparison.OrdinalIgnoreCase));
            if (!string.IsNullOrWhiteSpace(filePath))
            {
                var wanted = PathKey(filePath);
                candidates = candidates.Where(handle => string.Equals(PathKey(SafeString(() => ((dynamic)handle.Document).FullName)), wanted, StringComparison.OrdinalIgnoreCase));
            }
            else if (!string.IsNullOrWhiteSpace(name))
                candidates = candidates.Where(handle => string.Equals(SafeString(() => ((dynamic)handle.Document).Name), name, StringComparison.OrdinalIgnoreCase));
            else if (index > 0)
                candidates = candidates.Where(handle => SafeInt32(() => ((dynamic)handle.Document).Index) == index);
            else if (string.IsNullOrWhiteSpace(instanceId))
                throw new OfficeWorkerException("invalid_params", "需要 instanceId、filePath、name 或 index 之一");

            var matches = candidates.ToArray();
            if (matches.Length == 0)
                throw new OfficeWorkerException("document_not_found", "找不到指定实例和完整路径的 Office 文档窗口");
            if (matches.Length > 1)
                throw new OfficeWorkerException("ambiguous_document", "找到多个 Office 文档候选，请传 office.documents.list 返回的 instanceId 和完整路径");
            return new OfficeDocumentLease(matches[0], handles);
        }
        catch
        {
            ReleaseHandles(handles);
            throw;
        }
    }

    /// <summary>
    /// Smoke-only: detach Microsoft Excel Application RCW for a path+PID so the caller can Quit+Release.
    /// Releases all other document/application RCWs from the enumeration.
    /// </summary>
    internal static object? DetachExcelApplication(string fullPath, int processId)
    {
        if (processId <= 0) return null;
        var wanted = PathKey(fullPath);
        var handles = EnumerateDocuments("excel");
        object? ownedApp = null;
        try
        {
            var match = handles.FirstOrDefault(handle =>
                handle.ProcessId == processId
                && !OfficeHostRouting.IsWps(handle.ProgId)
                && string.Equals(
                    PathKey(SafeString(() => ((dynamic)handle.Document).FullName)),
                    wanted,
                    StringComparison.OrdinalIgnoreCase));
            if (match is null)
            {
                ReleaseHandles(handles);
                return null;
            }

            ownedApp = match.Application;
            var otherApps = new HashSet<object>(ReferenceEqualityComparer.Instance);
            foreach (var handle in handles)
            {
                ComInterop.Release(handle.Document);
                if (!ReferenceEquals(handle.Application, ownedApp))
                    otherApps.Add(handle.Application);
            }
            foreach (var application in otherApps)
                ComInterop.Release(application);
            return ownedApp;
        }
        catch
        {
            if (ownedApp is null) ReleaseHandles(handles);
            else
            {
                foreach (var handle in handles)
                {
                    ComInterop.Release(handle.Document);
                    if (!ReferenceEquals(handle.Application, ownedApp))
                        ComInterop.Release(handle.Application);
                }
            }
            throw;
        }
    }

    internal static string EncodeLocator(string value) => Uri.EscapeDataString(value);
    internal static string DecodeLocator(string value) => Uri.UnescapeDataString(value);
    internal static string PathKey(string? filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath)) return string.Empty;
        try { return Path.GetFullPath(filePath).ToLowerInvariant(); }
        catch { return filePath.ToLowerInvariant(); }
    }

    internal static string SafeString(Func<object?> value, string fallback = "")
    {
        try { return Convert.ToString(value(), System.Globalization.CultureInfo.InvariantCulture) ?? fallback; }
        catch { return fallback; }
    }

    internal static int SafeInt32(Func<object?> value, int fallback = 0)
    {
        try { return Convert.ToInt32(value(), System.Globalization.CultureInfo.InvariantCulture); }
        catch { return fallback; }
    }

    internal static long SafeInt64(Func<object?> value, long fallback = 0)
    {
        try { return Convert.ToInt64(value(), System.Globalization.CultureInfo.InvariantCulture); }
        catch { return fallback; }
    }

    internal static bool SafeBoolean(Func<object?> value, bool fallback)
    {
        try { return Convert.ToBoolean(value(), System.Globalization.CultureInfo.InvariantCulture); }
        catch { return fallback; }
    }

    private static List<OfficeDocumentHandle> EnumerateDocuments(string? appFilter)
    {
        var handles = new List<OfficeDocumentHandle>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in RunningObjectTable.Enumerate())
        {
            var kind = DocumentKind(entry.Value);
            if (string.IsNullOrWhiteSpace(kind) || (!string.IsNullOrWhiteSpace(appFilter) && kind != appFilter))
            {
                ComInterop.Release(entry.Value);
                continue;
            }
            AddHandle(handles, seen, entry.Value, kind, null, entry.DisplayName);
        }

        foreach (var config in ApplicationConfigs)
        {
            if (!string.IsNullOrWhiteSpace(appFilter) && config.App != appFilter) continue;
            foreach (var progId in config.ProgIds)
            {
                var application = ComInterop.TryGetActiveObject(progId);
                if (application is null) continue;
                object? collection = null;
                try
                {
                    collection = application.GetType().InvokeMember(config.Collection, System.Reflection.BindingFlags.GetProperty, null, application, null);
                    if (collection is null) continue;
                    dynamic items = collection;
                    var count = SafeInt32(() => items.Count);
                    for (var index = 1; index <= count; index++)
                    {
                        object? document = null;
                        try { document = items.Item(index); AddHandle(handles, seen, document, config.App, progId, string.Empty); }
                        catch { ComInterop.Release(document); }
                    }
                }
                catch { }
                finally
                {
                    ComInterop.Release(collection);
                    if (!handles.Any(handle => ReferenceEquals(handle.Application, application))) ComInterop.Release(application);
                }
            }
        }
        return handles.OrderBy(handle => handle.App).ThenBy(handle => handle.ProcessId).ThenBy(handle => handle.Hwnd).ThenBy(handle => SafeInt32(() => ((dynamic)handle.Document).Index)).ToList();
    }

    private static void AddHandle(List<OfficeDocumentHandle> handles, HashSet<string> seen, object? document, string app, string? progId, string rotName)
    {
        if (document is null) return;
        var handle = CreateHandle(document, app, progId, rotName);
        if (handle is null) { ComInterop.Release(document); return; }
        var fullName = PathKey(SafeString(() => ((dynamic)document).FullName));
        var index = SafeInt32(() => ((dynamic)document).Index);
        if (rotName.Length == 0 && fullName.Length > 0 && handles.Any(existing =>
                existing.App == app && PathKey(SafeString(() => ((dynamic)existing.Document).FullName)) == fullName))
        {
            ComInterop.Release(handle.Application);
            ComInterop.Release(document);
            return;
        }
        var identity = ComIdentity(document);
        var key = identity != 0 ? $"{app}|com:{identity:x}" : $"{handle.InstanceId}|{app}|{fullName}|{index}";
        if (seen.Add(key)) handles.Add(handle);
        else { ComInterop.Release(handle.Application); ComInterop.Release(document); }
    }

    private static OfficeDocumentHandle? CreateHandle(object document, string app, string? candidateProgId, string rotName)
    {
        object? application;
        try { application = ((dynamic)document).Application; }
        catch { return null; }
        if (application is null) return null;
        var documentName = SafeString(() => ((dynamic)document).Name);
        var applicationName = SafeString(() => ((dynamic)application).Name);
        if (string.IsNullOrWhiteSpace(documentName) || string.IsNullOrWhiteSpace(applicationName))
        {
            ComInterop.Release(application);
            return null;
        }
        var hwnd = ApplicationHwnd(application, document, app);
        var processId = WindowProcessId(hwnd);
        var processName = processId > 0 ? SafeString(() => Process.GetProcessById(processId).ProcessName) : string.Empty;
        var host = WpsHost(processName, candidateProgId, applicationName) ? "wps"
            : MicrosoftHost(processName, candidateProgId, applicationName) ? "microsoft-office" : "unknown";
        var progId = !string.IsNullOrWhiteSpace(candidateProgId) ? candidateProgId
            : host == "microsoft-office" ? app switch { "excel" => "Excel.Application", "word" => "Word.Application", _ => "PowerPoint.Application" }
            : host == "wps" ? app switch { "excel" => "Ket.Application", "word" => "Wps.Application", _ => "Wpp.Application" }
            : "ROT";
        var instanceId = processId > 0 || hwnd != 0
            ? $"{app}:{processId}:{hwnd}"
            : $"{app}:rot:{Uri.EscapeDataString(rotName)}";
        return new OfficeDocumentHandle(app, document, application, instanceId, processId, hwnd, host, progId);
    }

    private static long ComIdentity(object value)
    {
        if (!Marshal.IsComObject(value)) return 0;
        var pointer = IntPtr.Zero;
        try { pointer = Marshal.GetIUnknownForObject(value); return pointer.ToInt64(); }
        catch { return 0; }
        finally { if (pointer != IntPtr.Zero) Marshal.Release(pointer); }
    }

    private static object Metadata(OfficeDocumentHandle handle, bool? activeOverride = null)
    {
        dynamic document = handle.Document;
        return new
        {
            app = handle.App,
            name = SafeString(() => document.Name),
            fullName = SafeString(() => document.FullName),
            index = SafeInt32(() => document.Index),
            active = activeOverride ?? IsActive(handle),
            progId = handle.ProgId,
            host = handle.Host,
            instanceId = handle.InstanceId,
            processId = handle.ProcessId,
            hwnd = handle.Hwnd,
            readOnly = SafeBoolean(() => document.ReadOnly, false),
            saved = SafeBoolean(() => document.Saved, true),
        };
    }

    private static bool IsActive(OfficeDocumentHandle handle)
    {
        dynamic app = handle.Application;
        var activePath = handle.App switch
        {
            "excel" => SafeString(() => app.ActiveWorkbook?.FullName),
            "word" => SafeString(() => app.ActiveDocument?.FullName),
            _ => SafeString(() => app.ActivePresentation?.FullName),
        };
        return !string.IsNullOrWhiteSpace(activePath) && string.Equals(PathKey(activePath), PathKey(SafeString(() => ((dynamic)handle.Document).FullName)), StringComparison.OrdinalIgnoreCase);
    }

    private static string DocumentKind(object document)
    {
        try { _ = ((dynamic)document).Saved; }
        catch { return string.Empty; }
        var extension = Path.GetExtension(SafeString(() => ((dynamic)document).FullName)).ToLowerInvariant();
        if (extension is ".xlsx" or ".xlsm" or ".xlsb" or ".xls" or ".xltx" or ".xltm" or ".et") return "excel";
        if (extension is ".docx" or ".docm" or ".doc" or ".dotx" or ".dotm" or ".wps") return "word";
        if (extension is ".pptx" or ".pptm" or ".ppt" or ".potx" or ".potm" or ".dps") return "presentation";
        try { _ = ((dynamic)document).Worksheets; return "excel"; } catch { }
        try { _ = ((dynamic)document).Content; _ = ((dynamic)document).Bookmarks; return "word"; } catch { }
        try { _ = ((dynamic)document).Slides; return "presentation"; } catch { }
        return string.Empty;
    }

    private static long ApplicationHwnd(object application, object document, string app)
    {
        var hwnd = SafeInt64(() => ((dynamic)application).Hwnd);
        if (hwnd == 0) hwnd = SafeInt64(() => ((dynamic)application).HWND);
        if (hwnd != 0) return hwnd;
        return app switch
        {
            "excel" => SafeInt64(() => ((dynamic)document).Windows.Item(1).Hwnd),
            "word" => SafeInt64(() => ((dynamic)document).ActiveWindow.Hwnd, SafeInt64(() => ((dynamic)document).Windows.Item(1).Hwnd)),
            _ => SafeInt64(() => ((dynamic)document).Windows.Item(1).HWND, SafeInt64(() => ((dynamic)application).ActiveWindow.HWND)),
        };
    }

    private static int WindowProcessId(long hwnd)
    {
        if (hwnd == 0) return 0;
        _ = GetWindowThreadProcessId(new IntPtr(hwnd), out var processId);
        return unchecked((int)processId);
    }

    private static bool MicrosoftHost(string process, string? progId, string applicationName) =>
        process is "EXCEL" or "WINWORD" or "POWERPNT"
        || progId?.StartsWith("Excel.", StringComparison.OrdinalIgnoreCase) == true
        || progId?.StartsWith("Word.", StringComparison.OrdinalIgnoreCase) == true
        || progId?.StartsWith("PowerPoint.", StringComparison.OrdinalIgnoreCase) == true
        || applicationName.Contains("Microsoft", StringComparison.OrdinalIgnoreCase)
        || applicationName.Contains("Excel", StringComparison.OrdinalIgnoreCase)
        || applicationName.Contains("Word", StringComparison.OrdinalIgnoreCase)
        || applicationName.Contains("PowerPoint", StringComparison.OrdinalIgnoreCase);

    private static bool WpsHost(string process, string? progId, string applicationName) =>
        process is "wps" or "et" or "wpp" or "kso"
        || progId?.Contains("wps", StringComparison.OrdinalIgnoreCase) == true
        || progId?.Contains("ket", StringComparison.OrdinalIgnoreCase) == true
        || applicationName.Contains("WPS", StringComparison.OrdinalIgnoreCase)
        || applicationName.Contains("Kingsoft", StringComparison.OrdinalIgnoreCase);

    private static void PublishRestoreFiles(IReadOnlyList<RestoreFile> files)
    {
        var staged = new List<StagedRestore>();
        var committed = new List<StagedRestore>();
        try
        {
            foreach (var file in files)
            {
                var destination = Path.GetFullPath(file.FilePath);
                Directory.CreateDirectory(Path.GetDirectoryName(destination) ?? Environment.CurrentDirectory);
                string? stagedPath = null;
                if (file.Existed)
                {
                    if (string.IsNullOrWhiteSpace(file.SnapshotPath) || !File.Exists(file.SnapshotPath))
                        throw new OfficeWorkerException("snapshot_not_found", $"Office 事务快照不存在: {file.SnapshotPath}");
                    stagedPath = $"{destination}.{Guid.NewGuid():N}.transaction.stage";
                    File.Copy(file.SnapshotPath, stagedPath, true);
                }
                staged.Add(new StagedRestore(destination, stagedPath));
            }
            foreach (var entry in staged)
            {
                if (File.Exists(entry.Destination))
                {
                    entry.RollbackPath = $"{entry.Destination}.{Guid.NewGuid():N}.transaction.rollback";
                    File.Move(entry.Destination, entry.RollbackPath);
                }
                committed.Add(entry);
                if (!string.IsNullOrWhiteSpace(entry.StagedPath))
                {
                    File.Move(entry.StagedPath, entry.Destination);
                    entry.StagedPath = null;
                }
            }
            foreach (var entry in staged)
                if (!string.IsNullOrWhiteSpace(entry.RollbackPath) && File.Exists(entry.RollbackPath)) File.Delete(entry.RollbackPath);
        }
        catch
        {
            for (var index = committed.Count - 1; index >= 0; index--)
            {
                var entry = committed[index];
                if (File.Exists(entry.Destination)) File.Delete(entry.Destination);
                if (!string.IsNullOrWhiteSpace(entry.RollbackPath) && File.Exists(entry.RollbackPath))
                    File.Move(entry.RollbackPath, entry.Destination);
            }
            throw;
        }
        finally
        {
            foreach (var entry in staged)
            {
                if (!string.IsNullOrWhiteSpace(entry.StagedPath) && File.Exists(entry.StagedPath)) File.Delete(entry.StagedPath);
                if (!string.IsNullOrWhiteSpace(entry.RollbackPath) && File.Exists(entry.RollbackPath)) File.Delete(entry.RollbackPath);
            }
        }
    }

    private static IReadOnlyList<string> ReadStringArray(JsonElement value)
    {
        if (value.ValueKind != JsonValueKind.Array) return [];
        return value.EnumerateArray().Where(item => item.ValueKind == JsonValueKind.String).Select(item => item.GetString()).Where(item => !string.IsNullOrWhiteSpace(item)).Cast<string>().ToArray();
    }

    private static IReadOnlyList<RestoreFile> ReadRestoreFiles(JsonElement value)
    {
        if (value.ValueKind != JsonValueKind.Array) return [];
        var files = new List<RestoreFile>();
        foreach (var item in value.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.Object) continue;
            var filePath = item.OptionalString("filePath");
            if (string.IsNullOrWhiteSpace(filePath)) throw new OfficeWorkerException("invalid_params", "事务恢复文件缺少 filePath");
            files.Add(new RestoreFile(filePath, item.OptionalBoolean("existed"), item.OptionalString("snapshotPath")));
        }
        return files;
    }

    private static void ValidateApp(string? app, bool allowEmpty = false)
    {
        if (allowEmpty && string.IsNullOrWhiteSpace(app)) return;
        if (app is not ("excel" or "word" or "presentation"))
            throw new OfficeWorkerException("unsupported_app", $"不支持的 Office 应用: {app}");
    }

    internal static void ReleaseHandles(IEnumerable<OfficeDocumentHandle> handles)
    {
        var applications = new HashSet<object>(ReferenceEqualityComparer.Instance);
        foreach (var handle in handles)
        {
            applications.Add(handle.Application);
            ComInterop.Release(handle.Document);
        }
        foreach (var application in applications) ComInterop.Release(application);
    }

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

    private sealed record OfficeApplicationConfig(string App, string[] ProgIds, string Collection);
    private sealed record RestoreFile(string FilePath, bool Existed, string? SnapshotPath);
    private sealed record OpenDocumentSession(string App, string FilePath, string InstanceId, object Application, bool Active, bool ReadOnly);
    private sealed class StagedRestore(string destination, string? stagedPath)
    {
        public string Destination { get; } = destination;
        public string? StagedPath { get; set; } = stagedPath;
        public string? RollbackPath { get; set; }
    }

    private static class RunningObjectTable
    {
        public static IReadOnlyList<RotEntry> Enumerate()
        {
            var result = new List<RotEntry>();
            IRunningObjectTable? table = null;
            IBindCtx? context = null;
            IEnumMoniker? iterator = null;
            try
            {
                if (GetRunningObjectTable(0, out table) != 0 || table is null) return result;
                if (CreateBindCtx(0, out context) != 0 || context is null) return result;
                table.EnumRunning(out iterator);
                if (iterator is null) return result;
                iterator.Reset();
                var monikers = new IMoniker[1];
                while (iterator.Next(1, monikers, IntPtr.Zero) == 0)
                {
                    var moniker = monikers[0];
                    try
                    {
                        moniker.GetDisplayName(context, null, out var displayName);
                        table.GetObject(moniker, out var value);
                        if (value is not null) result.Add(new RotEntry(displayName ?? string.Empty, value));
                    }
                    catch { }
                    finally { if (moniker is not null && Marshal.IsComObject(moniker)) Marshal.ReleaseComObject(moniker); }
                }
                return result;
            }
            finally
            {
                if (iterator is not null && Marshal.IsComObject(iterator)) Marshal.ReleaseComObject(iterator);
                if (context is not null && Marshal.IsComObject(context)) Marshal.ReleaseComObject(context);
                if (table is not null && Marshal.IsComObject(table)) Marshal.ReleaseComObject(table);
            }
        }

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable runningObjectTable);
        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(int reserved, out IBindCtx bindContext);
    }

    private sealed record RotEntry(string DisplayName, object Value);
}

internal sealed record OfficeDocumentHandle(
    string App,
    object Document,
    object Application,
    string InstanceId,
    int ProcessId,
    long Hwnd,
    string Host,
    string ProgId);

internal sealed class OfficeDocumentLease(OfficeDocumentHandle handle, IReadOnlyList<OfficeDocumentHandle> handles) : IDisposable
{
    private bool disposed;

    public OfficeDocumentHandle Handle { get; } = handle;

    public void Dispose()
    {
        if (disposed) return;
        disposed = true;
        OfficeDocumentService.ReleaseHandles(handles);
    }
}
