using System.Diagnostics;
using Microsoft.Win32;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Office;

internal sealed class WpsDocumentLaunch(
    OfficeDocumentLease lease,
    bool ownsApplication,
    int launcherProcessId,
    IReadOnlySet<int> startedProcessIds) : IDisposable
{
    private bool disposed;

    public OfficeDocumentLease Lease { get; } = lease;
    public bool OwnsApplication { get; } = ownsApplication;

    public void Dispose()
    {
        if (disposed) return;
        disposed = true;
        Lease.Dispose();
        if (OwnsApplication) WpsFileLauncher.CleanupStartedProcesses(launcherProcessId, startedProcessIds);
    }
}

internal static class WpsFileLauncher
{
    public static WpsDocumentLaunch Open(string app, string filePath)
    {
        var fullPath = Path.GetFullPath(filePath);
        if (!File.Exists(fullPath)) throw new OfficeWorkerException("file_not_found", $"WPS 打开文件不存在: {fullPath}");
        var before = WpsProcessIds();
        var executable = ResolveExecutable(app)
            ?? throw new OfficeWorkerException("office_unavailable", "未找到 WPS Office 聚合启动程序");
        var component = app switch { "excel" => "/et", "word" => "/wps", _ => "/wpp" };
        int launcherProcessId;
        using (var process = Process.Start(new ProcessStartInfo(executable)
        {
            UseShellExecute = false,
            WindowStyle = ProcessWindowStyle.Minimized,
            CreateNoWindow = false,
            ArgumentList = { "/prometheus", component, fullPath },
        }))
        {
            if (process is null) throw new OfficeWorkerException("office_unavailable", $"无法启动 WPS 打开文件: {fullPath}");
            launcherProcessId = process.Id;
        }

        Exception? lastError = null;
        for (var attempt = 0; attempt < 100; attempt++)
        {
            try
            {
                var lease = OfficeDocumentService.AcquireDocument(app, fullPath, null);
                if (!OfficeHostRouting.IsWps(lease.Handle.ProgId))
                {
                    lease.Dispose();
                    throw new OfficeWorkerException("office_host_mismatch", $"文件未由 WPS 打开: {fullPath}");
                }
                var startedProcessIds = WpsProcessIds();
                startedProcessIds.ExceptWith(before);
                var ownsApplication = lease.Handle.ProcessId > 0 && !before.Contains(lease.Handle.ProcessId)
                    || StartedTargetComponent(app, before);
                return new WpsDocumentLaunch(
                    lease,
                    ownsApplication,
                    launcherProcessId,
                    startedProcessIds);
            }
            catch (OfficeWorkerException exception) when (exception.Code is "document_not_found" or "office_host_mismatch")
            {
                lastError = exception;
                Thread.Sleep(100);
            }
        }
        throw new OfficeWorkerException("office_open_timeout", $"WPS 未在超时前打开文件: {fullPath}", null, lastError);
    }

    private static string? ResolveExecutable(string app)
    {
        var executableName = app switch { "excel" => "et.exe", "word" => "wps.exe", _ => "wpp.exe" };
        var subKey = $"Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\{executableName}";
        foreach (var root in new[] { Registry.CurrentUser, Registry.LocalMachine })
        {
            using var key = root.OpenSubKey(subKey);
            var value = Convert.ToString(key?.GetValue(null));
            if (!string.IsNullOrWhiteSpace(value) && File.Exists(value)) return value;
        }
        return null;
    }

    private static HashSet<int> WpsProcessIds()
    {
        var result = new HashSet<int>();
        foreach (var name in new[] { "wps", "et", "wpp", "kso" })
            foreach (var process in Process.GetProcessesByName(name))
                using (process) result.Add(process.Id);
        return result;
    }

    private static bool StartedTargetComponent(string app, IReadOnlySet<int> existingProcessIds)
    {
        var processName = app switch { "excel" => "et", "word" => "wps", _ => "wpp" };
        foreach (var process in Process.GetProcessesByName(processName))
        {
            using (process)
                if (!existingProcessIds.Contains(process.Id)) return true;
        }
        return false;
    }

    internal static void CleanupStartedProcesses(int launcherProcessId, IReadOnlySet<int> startedProcessIds)
    {
        var ownedProcessIds = startedProcessIds.Append(launcherProcessId).Distinct().ToArray();
        for (var attempt = 0; attempt < 30; attempt++)
        {
            var remaining = ownedProcessIds.Where(ProcessExists).ToArray();
            if (remaining.Length == 0) return;
            Thread.Sleep(100);
        }

        foreach (var processId in ownedProcessIds.Where(ProcessExists))
        {
            try
            {
                using var process = Process.GetProcessById(processId);
                process.Kill(entireProcessTree: true);
                _ = process.WaitForExit(2_000);
            }
            catch { }
        }
    }

    private static bool ProcessExists(int processId)
    {
        try
        {
            using var process = Process.GetProcessById(processId);
            return !process.HasExited;
        }
        catch { return false; }
    }
}
