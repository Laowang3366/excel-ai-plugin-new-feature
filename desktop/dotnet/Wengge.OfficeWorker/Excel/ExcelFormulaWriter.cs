using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Excel;

/// <summary>
/// 共享、能力驱动的公式写入器。COM 与 Open XML 各自实现 <see cref="IFormulaTargetAdapter"/>
/// 适配宿主 API；写入器本身只关心分类、意图与回读校验。
///
/// 关键约束：
/// 1. 公式写入统一使用 <c>Range.Formula</c>，避免依赖不同宿主对 Formula2 的不稳定实现；
/// 2. <c>FormulaArray</c> 只走显式 CSE 意图（<see cref="FormulaWriteIntent.LegacyArray"/>）；
/// 3. 每个调用只写一个公式单元格，并做写后回读校验。
/// </summary>
internal interface IExcelFormulaCell
{
    void SetFormula(string formula);
    void SetFormulaArray(string formula);
    string? ReadFormula();
}

internal sealed class ComExcelFormulaCell(object cell, bool readFormulaArray = false) : IExcelFormulaCell
{
    private dynamic Api => cell;
    public void SetFormula(string formula) => Api.Formula = formula;
    public void SetFormulaArray(string formula) => Api.FormulaArray = formula;
    public string? ReadFormula() => Convert.ToString(readFormulaArray ? Api.FormulaArray : Api.Formula);
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
            if (kind == ExcelFormulaKind.LegacyArray) cell.SetFormulaArray(formula);
            else cell.SetFormula(formula);
        }
        catch (Exception ex)
        {
            var code = kind == ExcelFormulaKind.LegacyArray
                ? "legacy_array_unsupported"
                : "formula_rejected";
            var detail = new
            {
                exceptionType = ex.GetType().FullName,
                hresult = $"0x{ex.HResult:X8}",
                message = ex.Message,
            };
            throw new OfficeWorkerException(
                code,
                $"宿主拒绝了这条公式：{ex.Message}",
                detail,
                ex);
        }

        string? readBack;
        try
        {
            readBack = cell.ReadFormula();
        }
        catch (Exception ex)
        {
            throw new OfficeWorkerException("formula_verification_failed", "公式已写入但无法读回验证", null, ex);
        }
        if (string.IsNullOrWhiteSpace(readBack))
            throw new OfficeWorkerException("formula_verification_failed", $"{SetterName(kind)} 写后读回不符合预期");
        return new ExcelFormulaWriteResult(kind, readBack);
    }

    private static string SetterName(ExcelFormulaKind kind) => kind switch
    {
        ExcelFormulaKind.LegacyArray => "FormulaArray",
        _ => "Formula",
    };
}
