using System.Runtime.InteropServices;

namespace Wengge.OfficeWorker.Runtime;

internal sealed class ComMessageFilter : IOleMessageFilter
{
    private const int ServerCallRetryLater = 2;
    private const int MaxRetryElapsedMilliseconds = 5_000;

    public static void Register()
    {
        var filter = new ComMessageFilter();
        _ = CoRegisterMessageFilter(filter, out _);
    }

    public static void Revoke()
    {
        _ = CoRegisterMessageFilter(null, out _);
    }

    int IOleMessageFilter.HandleInComingCall(int callType, IntPtr taskCaller, int tickCount, IntPtr interfaceInfo) => 0;

    int IOleMessageFilter.RetryRejectedCall(IntPtr taskCallee, int tickCount, int rejectType) =>
        rejectType == ServerCallRetryLater && tickCount < MaxRetryElapsedMilliseconds ? 200 : -1;

    int IOleMessageFilter.MessagePending(IntPtr taskCallee, int tickCount, int pendingType) => 2;

    [DllImport("ole32.dll")]
    private static extern int CoRegisterMessageFilter(IOleMessageFilter? newFilter, out IOleMessageFilter? oldFilter);
}

[ComImport]
[Guid("00000016-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IOleMessageFilter
{
    [PreserveSig]
    int HandleInComingCall(int callType, IntPtr taskCaller, int tickCount, IntPtr interfaceInfo);

    [PreserveSig]
    int RetryRejectedCall(IntPtr taskCallee, int tickCount, int rejectType);

    [PreserveSig]
    int MessagePending(IntPtr taskCallee, int tickCount, int pendingType);
}
