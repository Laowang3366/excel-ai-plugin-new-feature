using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Excel;

/// <summary>
/// 共享、能力驱动的公式写入器。COM 与 Open XML 各自实现 <see cref="IFormulaTargetAdapter"/>
/// 适配宿主 API；写入器本身只关心分类、意图与回读校验。
///
/// 关键约束：
/// 1. 不得用 <c>Range.Formula</c> 静默回退把动态数组降级为 @；
/// 2. <c>FormulaArray</c> 只走显式 CSE 意图（<see cref="FormulaWriteIntent.LegacyArray"/>）；
/// 3. 每个调用只写一个公式单元格，并做写后回读校验。
/// </summary>
internal interface IExcelFormulaCell
{
    void SetFormula(string formula);
    void SetFormula2(string formula);
    void SetFormulaArray(string formula);
    string? ReadFormula();
    string? ReadFormula2();
}

internal sealed class ComExcelFormulaCell(object cell) : IExcelFormulaCell
{
    private dynamic Api => cell;
    public void SetFormula(string formula) => Api.Formula = formula;
    public void SetFormula2(string formula) => Api.Formula2 = formula;
    public void SetFormulaArray(string formula) => Api.FormulaArray = formula;
    public string? ReadFormula() => Convert.ToString(Api.Formula);
    public string? ReadFormula2() => Convert.ToString(Api.Formula2);
}

internal sealed record ExcelFormulaWriteResult(ExcelFormulaKind Kind, string? ReadBack);

internal static class ExcelFormulaWriter
{
    public static ExcelFormulaWriteResult Write(IExcelFormulaCell cell, string formula, bool legacyCse = false)
    {
        ArgumentNullException.ThrowIfNull(cell);
        var kind = ExcelFormulaClassification.Classify(formula, legacyCse);
        try
        {
            switch (kind)
            {
                case ExcelFormulaKind.Plain:
                    cell.SetFormula(formula);
                    break;
                case ExcelFormulaKind.Dynamic:
                    cell.SetFormula2(formula);
                    break;
                case ExcelFormulaKind.LegacyArray:
                    cell.SetFormulaArray(formula);
                    break;
            }
        }
        catch (Exception ex)
        {
            var code = kind switch
            {
                ExcelFormulaKind.Dynamic => "dynamic_array_unsupported",
                ExcelFormulaKind.LegacyArray => "legacy_array_unsupported",
                _ => "formula_write_failed",
            };
            throw new OfficeWorkerException(code, $"宿主无法通过 {SetterName(kind)} 写入公式", null, ex);
        }

        string? readBack;
        try
        {
            readBack = kind == ExcelFormulaKind.Dynamic ? cell.ReadFormula2() : cell.ReadFormula();
        }
        catch (Exception ex)
        {
            throw new OfficeWorkerException("formula_verification_failed", "公式已写入但无法读回验证", null, ex);
        }
        if (string.IsNullOrWhiteSpace(readBack) || kind == ExcelFormulaKind.Dynamic && readBack.StartsWith("=@", StringComparison.Ordinal))
            throw new OfficeWorkerException("formula_verification_failed", $"{SetterName(kind)} 写后读回不符合预期");
        return new ExcelFormulaWriteResult(kind, readBack);
    }

    private static string SetterName(ExcelFormulaKind kind) => kind switch
    {
        ExcelFormulaKind.Dynamic => "Formula2",
        ExcelFormulaKind.LegacyArray => "FormulaArray",
        _ => "Formula",
    };
}
