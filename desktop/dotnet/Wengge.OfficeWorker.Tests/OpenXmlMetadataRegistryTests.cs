using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using Wengge.OfficeWorker.OpenXml;

namespace Wengge.OfficeWorker.Tests;

public sealed class OpenXmlMetadataRegistryTests : IDisposable
{
    private const string MainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string DynamicArrayNamespace = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray";
    private readonly string root = Path.Combine(Path.GetTempPath(), $"wengge-metadata-{Guid.NewGuid():N}");

    public OpenXmlMetadataRegistryTests() => Directory.CreateDirectory(root);

    [Fact]
    public void RegisterDynamicArray_CreatesCellMetadataPartAndOneBasedIndex()
    {
        using var document = CreateDocument("metadata.xlsx");
        var workbookPart = document.WorkbookPart!;
        var registry = new OpenXmlMetadataRegistry();

        var index = registry.RegisterDynamicArray(workbookPart);

        Assert.Equal(1, index);
        Assert.NotNull(workbookPart.CellMetadataPart);
        Assert.StartsWith("/xl/metadata", workbookPart.CellMetadataPart!.Uri.ToString(), StringComparison.Ordinal);
        Assert.Equal("application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml",
            workbookPart.CellMetadataPart.ContentType);
        Assert.Equal("http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata",
            workbookPart.CellMetadataPart.RelationshipType);
        Assert.True(registry.IsDynamicArrayMetadata(workbookPart, index));
        var metadata = ReadXml(workbookPart.CellMetadataPart);
        Assert.Equal("metadata", metadata.Root!.Name.LocalName);
        Assert.Equal("XLDAPR", (string?)metadata.Root.Element(Name("metadataTypes"))!
            .Element(Name("metadataType"))!.Attribute("name"));
        var properties = metadata.Descendants(XName.Get("dynamicArrayProperties", DynamicArrayNamespace)).Single();
        Assert.Equal("1", (string?)properties.Attribute("fDynamic"));
        Assert.Equal("0", (string?)properties.Attribute("fCollapsed"));
        var record = metadata.Root.Element(Name("cellMetadata"))!.Element(Name("bk"))!.Element(Name("rc"))!;
        Assert.Equal("1", (string?)record.Attribute("t"));
        Assert.Equal("0", (string?)record.Attribute("v"));
    }

    [Fact]
    public void RegisterDynamicArray_ReusesExistingDescriptionAndCellMetadataBlock()
    {
        using var document = CreateDocument("reuse.xlsx");
        var workbookPart = document.WorkbookPart!;
        var registry = new OpenXmlMetadataRegistry();

        var first = registry.RegisterDynamicArray(workbookPart);
        var second = registry.RegisterDynamicArray(workbookPart);

        Assert.Equal(first, second);
        var metadata = ReadXml(workbookPart.CellMetadataPart!);
        Assert.Single(metadata.Root!.Elements(Name("futureMetadata"))
            .Single(element => (string?)element.Attribute("name") == "XLDAPR").Elements(Name("bk")));
        Assert.Single(metadata.Root.Element(Name("cellMetadata"))!.Elements(Name("bk")));
    }

    [Fact]
    public void RegisterDynamicArray_DoesNotWriteWorksheetExtension()
    {
        using var document = CreateDocument("worksheet.xlsx");
        var workbookPart = document.WorkbookPart!;
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new S.Worksheet(new S.SheetData());

        new OpenXmlMetadataRegistry().RegisterDynamicArray(workbookPart);

        Assert.DoesNotContain(
            worksheetPart.Worksheet.Descendants(),
            element => element.LocalName == "dynamicArrayProperties");
    }

    [Fact]
    public void RegisterDynamicArray_MergesWithExistingMetadataAndSurvivesReopen()
    {
        var path = Path.Combine(root, "merge.xlsx");
        using (var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new S.Workbook();
            var part = workbookPart.AddNewPart<CellMetadataPart>();
            var seed = new XDocument(
                new XElement(Name("metadata"),
                    new XElement(Name("metadataTypes"), new XAttribute("count", "1"),
                        new XElement(Name("metadataType"), new XAttribute("name", "EXISTING"))),
                    new XElement(Name("futureMetadata"), new XAttribute("name", "EXISTING"), new XAttribute("count", "1"),
                        new XElement(Name("bk"))),
                    new XElement(Name("cellMetadata"), new XAttribute("count", "1"),
                        new XElement(Name("bk"),
                            new XElement(Name("rc"), new XAttribute("t", "1"), new XAttribute("v", "0"))))));
            WriteXml(part, seed);

            Assert.Equal(2, new OpenXmlMetadataRegistry().RegisterDynamicArray(workbookPart));
            workbookPart.Workbook.Save();
        }

        using var reopened = SpreadsheetDocument.Open(path, false);
        var metadata = ReadXml(reopened.WorkbookPart!.CellMetadataPart!);
        Assert.Equal("2", (string?)metadata.Root!.Element(Name("metadataTypes"))!.Attribute("count"));
        Assert.Equal("2", (string?)metadata.Root.Element(Name("cellMetadata"))!.Attribute("count"));
        var record = metadata.Root.Element(Name("cellMetadata"))!.Elements(Name("bk")).Last().Element(Name("rc"))!;
        Assert.Equal("2", (string?)record.Attribute("t"));
        Assert.Equal("0", (string?)record.Attribute("v"));
    }

    private SpreadsheetDocument CreateDocument(string name)
    {
        var document = SpreadsheetDocument.Create(Path.Combine(root, name), SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new S.Workbook();
        return document;
    }

    private static XDocument ReadXml(OpenXmlPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        return XDocument.Load(stream);
    }

    private static void WriteXml(OpenXmlPart part, XDocument document)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        document.Save(stream);
    }

    private static XName Name(string localName) => XName.Get(localName, MainNamespace);

    public void Dispose()
    {
        try { Directory.Delete(root, recursive: true); } catch { }
    }
}
