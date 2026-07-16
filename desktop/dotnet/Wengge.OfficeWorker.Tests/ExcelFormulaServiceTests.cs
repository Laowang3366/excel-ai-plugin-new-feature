using Microsoft.CSharp.RuntimeBinder;
using Wengge.OfficeWorker.Excel;

namespace Wengge.OfficeWorker.Tests;

public sealed class ExcelFormulaServiceTests
{
    [Fact]
    public void TryGetFormulaText_AcceptsDynamicEqualsPrefixWithoutBinderException()
    {
        dynamic cell = "=B2:B4*2";
        Assert.True(ExcelFormulaService.TryGetFormulaText(cell, out string formula));
        Assert.Equal("=B2:B4*2", formula);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("90")]
    [InlineData("=")]
    public void TryGetFormulaText_RejectsNonFormulaCells(object? cell)
    {
        Assert.False(ExcelFormulaService.TryGetFormulaText(cell, out string formula));
        Assert.Equal(string.Empty, formula);
    }

    [Fact]
    public void DynamicStartsWithCharAndStringComparison_IsTheBinderFailureMode()
    {
        // Documents the production bug path: COM rows keep dynamic, Convert.ToString(dynamic)
        // stays dynamic, then StartsWith(char, StringComparison) fails to bind.
        dynamic raw = "=B2:B4";
        dynamic formula = Convert.ToString(raw);
        Assert.ThrowsAny<RuntimeBinderException>(() =>
            formula.StartsWith('=', StringComparison.Ordinal));
    }

    [Fact]
    public void DynamicArrayElementAtOrDefault_IsTheBinderFailureMode()
    {
        // Second real Excel failure: ToRows(dynamic) keeps the result dynamic, so LINQ
        // ElementAtOrDefault is resolved as an instance method on System.Array and fails.
        dynamic rows = new object?[][] { new object?[] { 90 } };
        Assert.ThrowsAny<RuntimeBinderException>(() => rows.ElementAtOrDefault(0));
    }

    [Fact]
    public void ToRows_OnStaticObject_SupportsElementAtOrDefault()
    {
        object? matrix = new object[,]
        {
            { 90 },
            { 80 },
        };
        object?[][] rows = ExcelValueConverter.ToRows(matrix);
        Assert.Equal(90, rows.ElementAtOrDefault(0)?.ElementAtOrDefault(0));
        Assert.Equal(80, rows.ElementAtOrDefault(1)?.ElementAtOrDefault(0));
    }
}
