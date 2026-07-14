using Wengge.OfficeWorker.Excel;

namespace Wengge.OfficeWorker.Tests;

public sealed class ExcelRangeWritePlanTests
{
    [Fact]
    public void Create_PureValuesRemainInBulkMatrix()
    {
        var matrix = new object[,] { { "Name", 1 }, { "Alice", 2 } };
        var plan = ExcelRangeWritePlan.Create(matrix);
        Assert.Empty(plan.Formulas);
        Assert.Equal("Alice", plan.BulkValues[1, 0]);
        Assert.Equal(0, plan.PlainCells);
    }

    [Fact]
    public void Create_MixedMatrixBlanksOnlyFormulaPositions()
    {
        var matrix = new object[,] { { 1, "=SUM(A1:A2)" }, { 2, "done" } };
        var plan = ExcelRangeWritePlan.Create(matrix);
        Assert.Single(plan.Formulas);
        Assert.Null(plan.BulkValues[0, 1]);
        Assert.Equal(2, plan.BulkValues[1, 0]);
        Assert.Equal("done", plan.BulkValues[1, 1]);
        Assert.Equal(1, plan.PlainCells);
    }

    [Fact]
    public void Create_PreservesEveryDynamicAnchor()
    {
        var matrix = new object[,] { { "=FILTER(A:A,A:A>0)", 7 }, { "=LET(x,1,x+1)", "tail" } };
        var plan = ExcelRangeWritePlan.Create(matrix);
        Assert.Equal(2, plan.Formulas.Count);
        Assert.Equal((0, 0), (plan.Formulas[0].Row, plan.Formulas[0].Column));
        Assert.Equal((1, 0), (plan.Formulas[1].Row, plan.Formulas[1].Column));
        Assert.Equal(2, plan.DynamicCells);
        Assert.Equal(7, plan.BulkValues[0, 1]);
        Assert.Equal("tail", plan.BulkValues[1, 1]);
    }

    [Fact]
    public void Create_LegacyCseClassifiesFormulaCellsAsArrays()
    {
        var plan = ExcelRangeWritePlan.Create(new object[,] { { "=SUM(A1:A10)" } }, legacyCse: true);
        Assert.Equal(1, plan.ArrayCells);
        Assert.Equal(ExcelFormulaKind.LegacyArray, plan.Formulas.Single().Kind);
    }
}
