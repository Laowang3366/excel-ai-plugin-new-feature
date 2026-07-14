using Wengge.OfficeWorker.Excel;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Tests;

public sealed class ExcelFormulaWriterTests
{
    [Fact]
    public void Write_UsesFormula2ForOrdinaryFormula()
    {
        var cell = new FakeFormulaCell();
        var result = ExcelFormulaWriter.Write(cell, "=SUM(A1:A10)");
        Assert.Equal("=SUM(A1:A10)", cell.Formula2);
        Assert.Null(cell.Formula);
        Assert.Equal(ExcelFormulaKind.Plain, result.Kind);
    }

    [Theory]
    [InlineData("=FILTER(A:A,A:A>0)")]
    [InlineData("=LET(x,1,x+1)")]
    public void Write_UsesFormula2ForModernFormula(string formula)
    {
        var cell = new FakeFormulaCell();
        var result = ExcelFormulaWriter.Write(cell, formula);
        Assert.Equal(formula, cell.Formula2);
        Assert.Null(cell.Formula);
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

    [Fact]
    public void Write_DoesNotFallbackWhenFormula2IsUnavailable()
    {
        var cell = new FakeFormulaCell { ThrowOnFormula2 = true };
        var error = Assert.Throws<OfficeWorkerException>(() =>
            ExcelFormulaWriter.Write(cell, "=FILTER(A:A,A:A>0)"));
        Assert.Equal("unsupported_formula_api", error.Code);
        Assert.Null(cell.Formula);
        Assert.Null(cell.FormulaArray);
    }

    [Theory]
    [InlineData("=A1:A3")]
    [InlineData("=A1:A3*2")]
    [InlineData("=IF(A1:A3>0,A1:A3,\"\")")]
    [InlineData("=TRANSPOSE(A1:A3)")]
    public void Write_UsesFormula2ForArrayExpressions(string formula)
    {
        var cell = new FakeFormulaCell();
        ExcelFormulaWriter.Write(cell, formula);
        Assert.Equal(formula, cell.Formula2);
        Assert.Null(cell.Formula);
    }

    [Fact]
    public void Write_RejectsImplicitIntersectionOnReadBack()
    {
        var cell = new FakeFormulaCell { Formula2ReadBack = "=@FILTER(A:A,A:A>0)" };
        var error = Assert.Throws<OfficeWorkerException>(() =>
            ExcelFormulaWriter.Write(cell, "=FILTER(A:A,A:A>0)"));
        Assert.Equal("formula_verification_failed", error.Code);
    }

    private sealed class FakeFormulaCell : IExcelFormulaCell
    {
        public string? Formula { get; private set; }
        public string? Formula2 { get; private set; }
        public string? FormulaArray { get; private set; }
        public bool ThrowOnFormula2 { get; init; }
        public string? Formula2ReadBack { get; init; }

        public void SetFormula(string formula) => Formula = formula;
        public void SetFormula2(string formula)
        {
            if (ThrowOnFormula2) throw new MissingMemberException("Formula2");
            Formula2 = formula;
        }
        public void SetFormulaArray(string formula) => FormulaArray = formula;
        public string? ReadFormula() => Formula ?? FormulaArray;
        public string? ReadFormula2() => Formula2ReadBack ?? Formula2;
    }
}
