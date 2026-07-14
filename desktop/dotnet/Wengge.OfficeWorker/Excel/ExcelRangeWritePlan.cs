namespace Wengge.OfficeWorker.Excel;

internal sealed record FormulaCellWrite(int Row, int Column, string Formula, ExcelFormulaKind Kind);

internal sealed record ExcelRangeWritePlan(
    object[,] BulkValues,
    IReadOnlyList<FormulaCellWrite> Formulas,
    int DynamicCells,
    int ArrayCells,
    int PlainCells)
{
    public static ExcelRangeWritePlan Create(object[,] matrix, bool legacyCse = false)
    {
        var rows = matrix.GetLength(0);
        var columns = matrix.GetLength(1);
        var bulkValues = (object[,])matrix.Clone();
        var formulas = new List<FormulaCellWrite>();
        var dynamicCells = 0;
        var arrayCells = 0;
        var plainCells = 0;

        for (var row = 0; row < rows; row++)
        {
            for (var column = 0; column < columns; column++)
            {
                if (matrix[row, column] is not string formula || !ExcelFormulaClassification.IsFormula(formula)) continue;
                var kind = ExcelFormulaClassification.Classify(formula, legacyCse);
                formulas.Add(new FormulaCellWrite(row, column, formula, kind));
                bulkValues[row, column] = null!;
                if (kind == ExcelFormulaKind.Dynamic) dynamicCells++;
                else if (kind == ExcelFormulaKind.LegacyArray) arrayCells++;
                else plainCells++;
            }
        }

        return new ExcelRangeWritePlan(
            bulkValues,
            formulas,
            dynamicCells,
            arrayCells,
            plainCells);
    }
}
