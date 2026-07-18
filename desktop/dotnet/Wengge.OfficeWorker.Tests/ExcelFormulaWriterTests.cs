using Wengge.OfficeWorker.Excel;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Tests;

public sealed class ExcelFormulaWriterTests
{
    [Fact]
    public void Write_UsesFormulaForOrdinaryFormula()
    {
        var cell = new FakeFormulaCell();
        var result = ExcelFormulaWriter.Write(cell, "=SUM(A1:A10)");
        Assert.Equal("=SUM(A1:A10)", cell.Formula);
        Assert.Null(cell.FormulaArray);
        Assert.Equal(ExcelFormulaKind.Plain, result.Kind);
    }

    [Theory]
    [InlineData("=FILTER(A:A,A:A>0)")]
    [InlineData("=LET(x,1,x+1)")]
    public void Write_UsesFormulaForModernFormula(string formula)
    {
        var cell = new FakeFormulaCell();
        var result = ExcelFormulaWriter.Write(cell, formula);
        Assert.Equal(formula, cell.Formula);
        Assert.Null(cell.FormulaArray);
        Assert.Equal(ExcelFormulaKind.Dynamic, result.Kind);
    }

    [Fact]
    public void Write_UsesFormulaArrayOnlyForExplicitLegacyCse()
    {
        var cell = new FakeFormulaCell();
        var result = ExcelFormulaWriter.Write(cell, "=SUM(A1:A10)", legacyCse: true);
        Assert.Equal("=SUM(A1:A10)", cell.FormulaArray);
        Assert.Equal(ExcelFormulaKind.LegacyArray, result.Kind);
    }

    [Theory]
    [InlineData("=A1:A3")]
    [InlineData("=A1:A3*2")]
    [InlineData("=IF(A1:A3>0,A1:A3,\"\")")]
    [InlineData("=TRANSPOSE(A1:A3)")]
    public void Write_UsesFormulaForArrayExpressions(string formula)
    {
        var cell = new FakeFormulaCell();
        ExcelFormulaWriter.Write(cell, formula);
        Assert.Equal(formula, cell.Formula);
        Assert.Null(cell.FormulaArray);
    }

    [Fact]
    public void Write_ReturnsHostReadBackWithoutRewritingIt()
    {
        var cell = new FakeFormulaCell { FormulaReadBack = "=@FILTER(A:A,A:A>0)" };
        var result = ExcelFormulaWriter.Write(cell, "=FILTER(A:A,A:A>0)");
        Assert.Equal("=@FILTER(A:A,A:A>0)", result.ReadBack);
    }

    private sealed class FakeFormulaCell : IExcelFormulaCell
    {
        public string? Formula { get; private set; }
        public string? FormulaArray { get; private set; }
        public string? FormulaReadBack { get; init; }

        public void SetFormula(string formula) => Formula = formula;
        public void SetFormulaArray(string formula) => FormulaArray = formula;
        public string? ReadFormula() => FormulaReadBack ?? Formula ?? FormulaArray;
    }
}
