using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Wengge.OfficeWorker.Com;

internal sealed class OwnedProcessJob : IDisposable
{
    private const uint JobObjectLimitKillOnJobClose = 0x00002000;
    private readonly IntPtr handle;
    private bool disposed;

    public OwnedProcessJob()
    {
        handle = CreateJobObject(IntPtr.Zero, null);
        if (handle == IntPtr.Zero) return;
        var limits = new JobObjectExtendedLimitInformation
        {
            BasicLimitInformation = new JobObjectBasicLimitInformation
            {
                LimitFlags = JobObjectLimitKillOnJobClose,
            },
        };
        var size = Marshal.SizeOf<JobObjectExtendedLimitInformation>();
        var buffer = Marshal.AllocHGlobal(size);
        try
        {
            Marshal.StructureToPtr(limits, buffer, false);
            if (!SetInformationJobObject(handle, 9, buffer, (uint)size))
            {
                CloseHandle(handle);
                this.handle = IntPtr.Zero;
            }
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }

    public HashSet<int> SnapshotCandidateProcessIds(string progId) =>
        CandidateProcessNames(progId)
            .SelectMany(Process.GetProcessesByName)
            .Select(process =>
            {
                using (process) return process.Id;
            })
            .ToHashSet();

    public OwnedApplicationProcess AssignCreatedApplication(object application, string progId, IReadOnlySet<int> existingProcesses)
    {
        try
        {
            dynamic app = application;
            var window = new IntPtr(Convert.ToInt64(app.Hwnd));
            if (window != IntPtr.Zero && GetWindowThreadProcessId(window, out var processId) != 0 && processId != 0)
            {
                if (existingProcesses.Contains((int)processId)) return new OwnedApplicationProcess((int)processId, false);
                if (handle != IntPtr.Zero && !disposed) _ = AssignProcess((int)processId);
                return new OwnedApplicationProcess((int)processId, true);
            }
        }
        catch
        {
            // Some WPS variants do not expose Hwnd until a document is visible.
        }

        for (var attempt = 0; attempt < 20; attempt++)
        {
            foreach (var processName in CandidateProcessNames(progId))
            {
                foreach (var process in Process.GetProcessesByName(processName))
                {
                    using (process)
                    {
                        if (existingProcesses.Contains(process.Id)) continue;
                        if (handle != IntPtr.Zero && !disposed) _ = AssignProcess(process.Id);
                        return new OwnedApplicationProcess(process.Id, true);
                    }
                }
            }
            Thread.Sleep(50);
        }
        return new OwnedApplicationProcess(null, false);
    }

    private bool AssignProcess(int processId)
    {
        try
        {
            using var process = Process.GetProcessById(processId);
            return AssignProcessToJobObject(handle, process.Handle);
        }
        catch { return false; }
    }

    private static string[] CandidateProcessNames(string progId) => progId.ToLowerInvariant() switch
    {
        "excel.application" => ["EXCEL"],
        "word.application" => ["WINWORD"],
        "powerpoint.application" => ["POWERPNT"],
        "ket.application" => ["et"],
        "wpp.application" or "kwpp.application" => ["wpp"],
        "kwps.application" or "wps.application" => ["wps"],
        _ => [],
    };

    public void Dispose()
    {
        if (disposed) return;
        disposed = true;
        if (handle != IntPtr.Zero) CloseHandle(handle);
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct JobObjectBasicLimitInformation
    {
        public long PerProcessUserTimeLimit;
        public long PerJobUserTimeLimit;
        public uint LimitFlags;
        public UIntPtr MinimumWorkingSetSize;
        public UIntPtr MaximumWorkingSetSize;
        public uint ActiveProcessLimit;
        public long Affinity;
        public uint PriorityClass;
        public uint SchedulingClass;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct IoCounters
    {
        public ulong ReadOperationCount;
        public ulong WriteOperationCount;
        public ulong OtherOperationCount;
        public ulong ReadTransferCount;
        public ulong WriteTransferCount;
        public ulong OtherTransferCount;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct JobObjectExtendedLimitInformation
    {
        public JobObjectBasicLimitInformation BasicLimitInformation;
        public IoCounters IoInfo;
        public UIntPtr ProcessMemoryLimit;
        public UIntPtr JobMemoryLimit;
        public UIntPtr PeakProcessMemoryUsed;
        public UIntPtr PeakJobMemoryUsed;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr CreateJobObject(IntPtr securityAttributes, string? name);

    [DllImport("kernel32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetInformationJobObject(IntPtr job, int informationClass, IntPtr information, uint informationLength);

    [DllImport("kernel32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

    [DllImport("kernel32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CloseHandle(IntPtr handle);
}

internal sealed record OwnedApplicationProcess(int? ProcessId, bool Owned);
