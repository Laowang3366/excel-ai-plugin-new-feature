using System.Runtime.ExceptionServices;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Excel;

internal static class ExcelRangeWriteTransaction
{
    public static void Execute<TSnapshot>(
        Func<TSnapshot> capture,
        Action write,
        Action<TSnapshot> restore)
    {
        var snapshot = capture();
        try
        {
            write();
        }
        catch (Exception writeError)
        {
            try
            {
                restore(snapshot);
            }
            catch (Exception rollbackError)
            {
                throw new OfficeWorkerException(
                    "range_write_rollback_failed",
                    "范围写入失败，且无法完整恢复原内容",
                    new { writeError = writeError.Message, rollbackError = rollbackError.Message },
                    rollbackError);
            }

            ExceptionDispatchInfo.Capture(writeError).Throw();
            throw;
        }
    }
}
