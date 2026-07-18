using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Com;

internal sealed class OfficeApplicationProvider(OwnedProcessJob ownedProcesses)
{
    public IReadOnlyList<string> DetectActiveProgIds(IEnumerable<string> progIds)
    {
        var active = new List<string>();
        foreach (var progId in progIds)
        {
            var application = ComInterop.TryGetActiveObject(progId);
            if (application is null) continue;
            active.Add(progId);
            ComInterop.Release(application);
        }

        return active;
    }

    public OfficeApplicationHandle? TryGetActive(IEnumerable<string> progIds)
    {
        foreach (var progId in progIds)
        {
            var application = ComInterop.TryGetActiveObject(progId);
            if (application is not null)
            {
                return new OfficeApplicationHandle(application, progId, created: false);
            }
        }

        return null;
    }

    public OfficeApplicationHandle GetActiveRequired(IEnumerable<string> progIds, string message)
    {
        return TryGetActive(progIds) ?? throw new OfficeWorkerException("office_not_connected", message);
    }

    public OfficeApplicationHandle GetOrCreate(IEnumerable<string> progIds, string message)
    {
        var ids = progIds.ToArray();
        var active = TryGetActive(ids);
        if (active is not null)
        {
            return active;
        }

        return Create(ids, message);
    }

    public OfficeApplicationHandle Create(IEnumerable<string> progIds, string message)
    {
        Exception? lastError = null;
        var ids = progIds.ToArray();
        foreach (var progId in ids)
        {
            object? application = null;
            try
            {
                var existingProcesses = ownedProcesses.SnapshotCandidateProcessIds(progId);
                application = ComInterop.CreateObject(progId);
                var process = ownedProcesses.AssignCreatedApplication(application, progId, existingProcesses);
                if (!process.Owned)
                {
                    throw new OfficeWorkerException(
                        "office_instance_not_isolated",
                        $"{progId} 未创建独立进程；为保护用户已打开的 Office 窗口，本次文件级操作已停止。请关闭已有窗口后重试，或显式传入 instanceId。");
                }
                var result = new OfficeApplicationHandle(application, progId, created: true, process.ProcessId);
                application = null;
                return result;
            }
            catch (Exception exception)
            {
                lastError = exception;
            }
            finally { ComInterop.Release(application); }
        }

        if (lastError is OfficeWorkerException workerError && workerError.Code == "office_instance_not_isolated") throw workerError;
        throw new OfficeWorkerException("office_unavailable", message, null, lastError);
    }
}

internal sealed class OfficeApplicationHandle(object application, string progId, bool created, int? processId = null) : IDisposable
{
    public dynamic Application { get; } = application;

    public string ProgId { get; } = progId;

    public bool Created { get; } = created;

    public int? ProcessId { get; } = processId;

    public void Dispose() => ComInterop.Release(Application);

    public bool WaitForExit(int timeoutMilliseconds = 10000)
    {
        if (!Created || ProcessId is not int id) return true;
        try
        {
            using var process = System.Diagnostics.Process.GetProcessById(id);
            return process.HasExited || process.WaitForExit(timeoutMilliseconds);
        }
        catch (ArgumentException) { return true; }
        catch { return false; }
    }
}
