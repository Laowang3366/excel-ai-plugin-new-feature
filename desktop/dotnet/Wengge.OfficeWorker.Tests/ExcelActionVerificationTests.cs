using System.Text.Json;
using Wengge.OfficeWorker.Office;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Tests;

public sealed class ExcelActionVerificationTests
{
    [Fact]
    public void VerifyInsertedChart_AcceptsVisibleChartWithSeriesAndAnchors()
    {
        var result = ExcelActionService.VerifyInsertedChart(
            new FakeChartObject(), new FakeChart(), new FakeCollection(2));
        Assert.True(result.Ok, JsonSerializer.Serialize(result.Checks));
        Assert.Equal("D2", result.TopLeftCell);
        Assert.Equal("J15", result.BottomRightCell);
        Assert.Equal(2, result.SeriesCount);
    }

    [Fact]
    public void VerifyInsertedChart_RejectsInvisibleOrEmptyChart()
    {
        var invisible = ExcelActionService.VerifyInsertedChart(
            new FakeChartObject { Visible = false }, new FakeChart(), new FakeCollection(2));
        var empty = ExcelActionService.VerifyInsertedChart(
            new FakeChartObject(), new FakeChart(), new FakeCollection(0));
        Assert.False(invisible.Ok);
        Assert.False(empty.Ok);
    }

    [Fact]
    public void VerifyInsertedChart_RejectsChartOnVeryHiddenSheet()
    {
        var result = ExcelActionService.VerifyInsertedChart(
            new FakeChartObject { Parent = new FakeWorksheet { Visible = 2 } },
            new FakeChart(),
            new FakeCollection(1));
        Assert.False(result.Ok);
        Assert.False(result.SheetVisible);
    }

    [Theory]
    [InlineData(null, true, null, "A1")]
    [InlineData("", true, null, "A1")]
    [InlineData("Summary!D5", false, "Summary", "D5")]
    [InlineData("'Pivot Data'!A3", false, "Pivot Data", "A3")]
    [InlineData("H3", false, null, "H3")]
    public void ParsePivotDestination_SeparatesDedicatedSheetAndExplicitTargets(
        string? value, bool dedicated, string? sheet, string address)
    {
        var result = ExcelActionService.ParsePivotDestination(value);
        Assert.Equal(dedicated, result.UseDedicatedSheet);
        Assert.Equal(sheet, result.SheetName);
        Assert.Equal(address, result.Address);
    }

    [Fact]
    public void VerifyPivotTable_RequiresRealCacheFieldsAndDestination()
    {
        ExcelActionService.PivotPostWrite result;
        try { result = ExcelActionService.VerifyPivotTable(new FakePivot(), "Pivot1", "Data!R1C1:R10C3"); }
        catch (OfficeWorkerException error)
        {
            Assert.Fail(JsonSerializer.Serialize(error.Details));
            throw;
        }
        Assert.Equal("A1:C5", result.DestinationRange);
        Assert.Equal(2, result.RowFieldCount + result.DataFieldCount);
    }

    [Fact]
    public void VerifyPivotTable_RejectsMissingCacheOrFields()
    {
        var missingCache = new FakePivot { Cache = null };
        var noFields = new FakePivot
        {
            RowFields = new FakeCollection(0),
            DataFields = new FakeCollection(0),
        };
        Assert.Equal("pivot_verification_failed",
            Assert.Throws<OfficeWorkerException>(() =>
                ExcelActionService.VerifyPivotTable(missingCache, "Pivot1", "Data!R1C1:R10C3")).Code);
        Assert.Equal("pivot_verification_failed",
            Assert.Throws<OfficeWorkerException>(() =>
                ExcelActionService.VerifyPivotTable(noFields, "Pivot1", "Data!R1C1:R10C3")).Code);
    }

    public sealed class FakeChartObject
    {
        public string Name { get; init; } = "Chart 1";
        public bool Visible { get; init; } = true;
        public FakeWorksheet Parent { get; init; } = new();
        public double Left { get; init; } = 120;
        public double Top { get; init; } = 20;
        public double Width { get; init; } = 420;
        public double Height { get; init; } = 260;
        public FakeRange TopLeftCell { get; } = new("D2");
        public FakeRange BottomRightCell { get; } = new("J15");
    }

    public sealed class FakeWorksheet
    {
        public string Name { get; init; } = "Data";
        public int Visible { get; init; } = -1;
    }

    public sealed class FakeChart
    {
        public bool HasTitle { get; init; }
        public FakeChartTitle ChartTitle { get; } = new();
    }

    public sealed class FakeChartTitle
    {
        public string Text { get; init; } = string.Empty;
    }

    public sealed class FakeCollection(int count)
    {
        public int Count { get; } = count;
    }

    public sealed class FakePivot
    {
        public FakeCollection RowFields { get; init; } = new(1);
        public FakeCollection ColumnFields { get; init; } = new(0);
        public FakeCollection PageFields { get; init; } = new(0);
        public FakeCollection DataFields { get; init; } = new(1);
        public object? Cache { get; init; } = new();
        public FakeRange TableRange1 { get; init; } = new("A1:C5");
        public FakeRange TableRange2 { get; init; } = new("A1:C6");
        public FakeRange? DataBodyRange { get; init; } = new("B2:C5");
        public object? PivotCache() => Cache;
    }

    public sealed class FakeRange(string address)
    {
        public string Address { get; } = address;
    }
}
