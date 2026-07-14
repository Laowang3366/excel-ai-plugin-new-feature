using Wengge.OfficeWorker.Excel;

namespace Wengge.OfficeWorker.Tests;

public sealed class ExcelFormulaClassificationTests
{
    [Theory]
    [InlineData(null, false)]
    [InlineData("", false)]
    [InlineData("text", false)]
    [InlineData("=", false)]
    [InlineData("=SUM(A1:A10)", true)]
    public void IsFormula_RequiresEqualsAndBody(string? value, bool expected) =>
        Assert.Equal(expected, ExcelFormulaClassification.IsFormula(value));

    [Theory]
    [InlineData("=FILTER(A:A,A:A<>\"\")")]
    [InlineData("=SORTBY(A1:A5,B1:B5)")]
    [InlineData("=LET(x,1,x+1)")]
    [InlineData("=LAMBDA(x,x+1)(1)")]
    [InlineData("=XLOOKUP(A1,B:B,C:C)")]
    [InlineData("=MAP(A1:A5,LAMBDA(x,x+1))")]
    [InlineData("=SUM(FILTER(A1:A10,A1:A10>0))")]
    [InlineData("=IF(A1,FILTER(B:B,B:B<>\"\"),\"\")")]
    [InlineData("=GROUPBY(A:A,B:B,SUM)")]
    [InlineData("=BYROW(A1:C3,LAMBDA(r,SUM(r)))")]
    [InlineData("=_xlfn._xlws.FILTER(A:A,A:A>0)")]
    public void Classify_RecognizesModernFormulaHead(string formula) =>
        Assert.Equal(ExcelFormulaKind.Dynamic, ExcelFormulaClassification.Classify(formula));

    [Theory]
    [InlineData("=SUM(A1:A10)")]
    [InlineData("=TODAY()")]
    [InlineData("=IF(A1>0,1,0)")]
    [InlineData("=\"FILTER(A:A)\"")]
    [InlineData("='FILTER(A)'!A1")]
    public void Classify_KeepsOrdinaryFormulaPlain(string formula) =>
        Assert.Equal(ExcelFormulaKind.Plain, ExcelFormulaClassification.Classify(formula));

    [Fact]
    public void Classify_UsesLegacyArrayOnlyWhenExplicitlyRequested() =>
        Assert.Equal(ExcelFormulaKind.LegacyArray, ExcelFormulaClassification.Classify("=SUM(A1:A10)", legacyCse: true));

    [Theory]
    [InlineData("=A1:A3")]
    [InlineData("=A1:A3*2")]
    [InlineData("=IF(A1:A3>0,A1:A3,\"\")")]
    [InlineData("=TRANSPOSE(A1:A3)")]
    [InlineData("=MYFILTER(A1:A10)")]
    public void IsDynamicArray_RecognizesExpressionBasedSpill(string formula) =>
        Assert.True(ExcelFormulaClassification.IsDynamicArray(formula));

    [Theory]
    [InlineData("=SUM(A1:A10)")]
    [InlineData("=IF(A1>0,1,0)")]
    [InlineData("=\"A1:A3\"")]
    public void IsDynamicArray_DoesNotPromoteScalarOrQuotedRanges(string formula) =>
        Assert.False(ExcelFormulaClassification.IsDynamicArray(formula));

    [Theory]
    [InlineData("=FILTER(A:A,A:A>0)", "_xlfn._xlws.FILTER(A:A,A:A>0)")]
    [InlineData("=LET(x,1,x+1)", "_xlfn.LET(x,1,x+1)")]
    [InlineData("=XLOOKUP(A1,B:B,C:C)", "_xlfn.XLOOKUP(A1,B:B,C:C)")]
    [InlineData("=GROUPBY(A:A,B:B,SUM)", "_xlfn._xlws.GROUPBY(A:A,B:B,SUM)")]
    [InlineData("=SUM(A1:A10)", "SUM(A1:A10)")]
    public void NormalizeForOpenXml_AddsExpectedPrefix(string formula, string expected) =>
        Assert.Equal(expected, ExcelFormulaClassification.NormalizeForOpenXml(formula));

    [Fact]
    public void NormalizeForOpenXml_DoesNotDoublePrefix() =>
        Assert.Equal("_xlfn._xlws.FILTER(A:A,A:A>0)",
            ExcelFormulaClassification.NormalizeForOpenXml("=_xlfn._xlws.FILTER(A:A,A:A>0)"));

    [Fact]
    public void NormalizeForOpenXml_PrefixesNestedFunctionsButNotStringLiterals()
    {
        Assert.Equal(
            "_xlfn.LET(data,'FILTER(A)'!A:A,_xlfn._xlws.FILTER(data,data>0))&\"FILTER(A:A)\"",
            ExcelFormulaClassification.NormalizeForOpenXml("=LET(data,'FILTER(A)'!A:A,FILTER(data,data>0))&\"FILTER(A:A)\""));
    }
}
