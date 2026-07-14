using Wengge.OfficeWorker.Excel;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Tests;

public sealed class ExcelRangeWriteTransactionTests
{
    [Fact]
    public void Execute_RestoresSnapshotWhenLaterFormulaFails()
    {
        var cells = new[] { "before-1", "before-2" };

        var error = Assert.Throws<InvalidOperationException>(() =>
            ExcelRangeWriteTransaction.Execute(
                () => cells.ToArray(),
                () =>
                {
                    cells[0] = "written";
                    cells[1] = "partial";
                    throw new InvalidOperationException("second Formula2 failed");
                },
                snapshot => Array.Copy(snapshot, cells, snapshot.Length)));

        Assert.Equal("second Formula2 failed", error.Message);
        Assert.Equal(new[] { "before-1", "before-2" }, cells);
    }

    [Fact]
    public void Execute_ReportsRollbackFailureExplicitly()
    {
        var error = Assert.Throws<OfficeWorkerException>(() =>
            ExcelRangeWriteTransaction.Execute(
                () => "snapshot",
                () => throw new InvalidOperationException("write failed"),
                _ => throw new InvalidOperationException("restore failed")));

        Assert.Equal("range_write_rollback_failed", error.Code);
    }
}
