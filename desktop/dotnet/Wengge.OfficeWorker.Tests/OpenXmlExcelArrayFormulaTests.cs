using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using Wengge.OfficeWorker.Protocol;
using Wengge.OfficeWorker.Runtime;

namespace Wengge.OfficeWorker.Tests;

public sealed class OpenXmlExcelArrayFormulaTests : IDisposable
{
    private readonly string root = Path.Combine(Path.GetTempPath(), $"wengge-array-{Guid.NewGuid():N}");

    public OpenXmlExcelArrayFormulaTests() => Directory.CreateDirectory(root);

    [Fact]
    public async Task SingleDynamicAnchor_WritesArrayFormulaAndCellMetadata()
    {
        var path = Path.Combine(root, "single.xlsx");
        var result = await CreateWorkbook(path, "range:Data!J20", [["=FILTER(A2:A100,B2:B100>0)"]]);

        Assert.Equal(1, result.GetProperty("data").GetProperty("dynamicAnchors").GetInt32());
        using var document = SpreadsheetDocument.Open(path, false);
        var workbookPart = document.WorkbookPart!;
        var cell = workbookPart.WorksheetParts.Single().Worksheet.Descendants<S.Cell>()
            .Single(candidate => candidate.CellReference == "J20");
        Assert.Equal(1U, cell.CellMetaIndex!.Value);
        Assert.Equal(S.CellFormulaValues.Array, cell.CellFormula!.FormulaType!.Value);
        Assert.Equal("J20", cell.CellFormula.Reference!.Value);
        Assert.Equal("_xlfn._xlws.FILTER(A2:A100,B2:B100>0)", cell.CellFormula.Text);
        Assert.NotNull(workbookPart.CellMetadataPart);
        Assert.Empty(workbookPart.WorksheetParts.Single().Worksheet.Elements<S.ExtensionList>());
    }

    [Fact]
    public async Task MixedMatrix_PreservesValuesOrdinaryFormulaAndEveryDynamicAnchor()
    {
        var path = Path.Combine(root, "mixed.xlsx");
        await CreateWorkbook(path, "range:Data!A1:B2",
        [
            ["=FILTER(D:D,D:D>0)", 42],
            ["=SUM(D1:D10)", "=LET(x,1,x+1)"],
        ]);

        using var document = SpreadsheetDocument.Open(path, false);
        var cells = document.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<S.Cell>()
            .ToDictionary(cell => cell.CellReference!.Value!);
        Assert.Equal(4, cells.Count);
        Assert.Equal("42", cells["B1"].CellValue!.Text);
        var staticFormula = cells["A2"].CellFormula!;
        Assert.Equal("SUM(D1:D10)", staticFormula.Text);
        Assert.Null(staticFormula.FormulaType);
        Assert.Equal(1U, cells["A1"].CellMetaIndex!.Value);
        Assert.Equal(1U, cells["B2"].CellMetaIndex!.Value);
        Assert.Equal(S.CellFormulaValues.Array, cells["A1"].CellFormula!.FormulaType!.Value);
        Assert.Equal(S.CellFormulaValues.Array, cells["B2"].CellFormula!.FormulaType!.Value);
    }

    [Fact]
    public async Task DynamicAnchor_OmitsEmptySpillPlaceholders()
    {
        var path = Path.Combine(root, "spill-placeholders.xlsx");
        await CreateWorkbook(path, "range:Data!A1:C1", [["=FILTER(D:D,D:D>0)", "", ""]]);

        using var document = SpreadsheetDocument.Open(path, false);
        var cells = document.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<S.Cell>().ToArray();
        var cell = Assert.Single(cells);
        Assert.Equal("A1", cell.CellReference!.Value);
        Assert.Equal(S.CellFormulaValues.Array, cell.CellFormula!.FormulaType!.Value);
    }

    [Fact]
    public async Task StaticFormula_DoesNotCreateDynamicMetadata()
    {
        var path = Path.Combine(root, "static.xlsx");
        await CreateWorkbook(path, "range:Data!A1", [["=TODAY()"]]);

        using var document = SpreadsheetDocument.Open(path, false);
        var cell = document.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<S.Cell>().Single();
        Assert.Null(document.WorkbookPart.CellMetadataPart);
        Assert.Null(cell.CellFormula!.FormulaType);
        Assert.Null(cell.CellMetaIndex);
    }

    private static async Task<JsonElement> CreateWorkbook(string path, string target, object[][] values)
    {
        using var worker = OfficeWorkerHost.Create();
        var parameters = JsonSerializer.SerializeToElement(new
        {
            app = "excel",
            action = "insert",
            operation = "createWorkbook",
            filePath = path,
            target,
            @params = new { sheetNames = new[] { "Data" }, values },
        }, new JsonSerializerOptions(JsonSerializerDefaults.Web));
        var response = await worker.DispatchAsync(
            new RpcRequest("test", "openxml.action.execute", parameters), CancellationToken.None);
        var result = JsonSerializer.SerializeToElement(response, new JsonSerializerOptions(JsonSerializerDefaults.Web));
        Assert.Equal("done", result.GetProperty("status").GetString());
        return result;
    }

    public void Dispose()
    {
        try { Directory.Delete(root, recursive: true); } catch { }
    }
}
