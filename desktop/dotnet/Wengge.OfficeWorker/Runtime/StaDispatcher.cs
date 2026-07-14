using System.Collections.Concurrent;
using System.Runtime.InteropServices;

namespace Wengge.OfficeWorker.Runtime;

public sealed class StaDispatcher : IDisposable
{
    private readonly BlockingCollection<IWorkItem> queue = new();
    private readonly AutoResetEvent queueSignal = new(false);
    private readonly Thread thread;
    private bool disposed;

    public StaDispatcher()
    {
        thread = new Thread(Run)
        {
            IsBackground = true,
            Name = "Wengge.OfficeWorker.COM",
        };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
    }

    public Task<T> InvokeAsync<T>(Func<T> operation, CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(disposed, this);
        var item = new WorkItem<T>(operation, cancellationToken);
        queue.Add(item, cancellationToken);
        queueSignal.Set();
        return item.Task;
    }

    public void Dispose()
    {
        if (disposed)
        {
            return;
        }

        disposed = true;
        queue.CompleteAdding();
        queueSignal.Set();
        if (thread.Join(TimeSpan.FromSeconds(5)))
        {
            queueSignal.Dispose();
            queue.Dispose();
        }
    }

    private void Run()
    {
        ComMessageFilter.Register();
        try
        {
            while (!queue.IsCompleted)
            {
                while (queue.TryTake(out var item)) item.Execute();
                if (queue.IsCompleted) break;
                WaitForWorkOrMessages();
            }
        }
        finally
        {
            ComMessageFilter.Revoke();
        }
    }

    private void WaitForWorkOrMessages()
    {
        var handles = new[] { queueSignal.SafeWaitHandle.DangerousGetHandle() };
        _ = MsgWaitForMultipleObjectsEx(1, handles, uint.MaxValue, QsAllInput, MwmoInputAvailable);
        while (PeekMessage(out var message, IntPtr.Zero, 0, 0, PmRemove))
        {
            TranslateMessage(ref message);
            DispatchMessage(ref message);
        }
    }

    private interface IWorkItem
    {
        void Execute();
    }

    private sealed class WorkItem<T>(Func<T> operation, CancellationToken cancellationToken) : IWorkItem
    {
        private readonly TaskCompletionSource<T> completion =
            new(TaskCreationOptions.RunContinuationsAsynchronously);

        public Task<T> Task => completion.Task;

        public void Execute()
        {
            if (cancellationToken.IsCancellationRequested)
            {
                completion.TrySetCanceled(cancellationToken);
                return;
            }

            try
            {
                completion.TrySetResult(operation());
            }
            catch (Exception exception)
            {
                completion.TrySetException(exception);
            }
        }
    }

    private const uint QsAllInput = 0x04FF;
    private const uint MwmoInputAvailable = 0x0004;
    private const uint PmRemove = 0x0001;

    [StructLayout(LayoutKind.Sequential)]
    private struct Message
    {
        public IntPtr Window;
        public uint Value;
        public UIntPtr WParam;
        public IntPtr LParam;
        public uint Time;
        public Point Point;
        public uint Private;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct Point
    {
        public int X;
        public int Y;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint MsgWaitForMultipleObjectsEx(uint count, IntPtr[] handles, uint milliseconds, uint wakeMask, uint flags);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool PeekMessage(out Message message, IntPtr window, uint minimum, uint maximum, uint remove);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool TranslateMessage(ref Message message);

    [DllImport("user32.dll")]
    private static extern IntPtr DispatchMessage(ref Message message);
}
