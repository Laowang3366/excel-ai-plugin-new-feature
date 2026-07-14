using System.Runtime.InteropServices;

namespace Wengge.OfficeWorker.Com;

internal static class ComInterop
{
    public static object? TryGetActiveObject(string progId)
    {
        var classId = Guid.Empty;
        var classResult = CLSIDFromProgID(progId, out classId);
        if (classResult < 0)
        {
            return null;
        }

        var result = GetActiveObject(ref classId, IntPtr.Zero, out var instance);
        return result >= 0 ? instance : null;
    }

    public static object CreateObject(string progId)
    {
        var type = Type.GetTypeFromProgID(progId, throwOnError: false)
            ?? throw new COMException($"没有注册 COM ProgID: {progId}");
        return Activator.CreateInstance(type)
            ?? throw new COMException($"无法创建 COM 对象: {progId}");
    }

    public static void Release(object? instance)
    {
        if (instance is null || !Marshal.IsComObject(instance))
        {
            return;
        }

        try
        {
            _ = Marshal.ReleaseComObject(instance);
        }
        catch (InvalidComObjectException)
        {
            // Another local scope already released this RCW.
        }
    }

    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgID(string progId, out Guid classId);

    [DllImport("oleaut32.dll")]
    private static extern int GetActiveObject(
        ref Guid classId,
        IntPtr reserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object? instance);
}
