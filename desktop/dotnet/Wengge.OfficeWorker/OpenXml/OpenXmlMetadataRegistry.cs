using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.OpenXml;

internal sealed class OpenXmlMetadataRegistry
{
    private const string MainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string DynamicArrayNamespace = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray";
    private const string DynamicArrayExtensionUri = "{BDBB8CDC-FA1E-496E-A857-3C3F30C029C3}";
    private readonly object gate = new();

    public int RegisterDynamicArray(WorkbookPart workbookPart)
    {
        ArgumentNullException.ThrowIfNull(workbookPart);
        lock (gate)
        {
            var part = EnsureMetadataPart(workbookPart);
            var document = ReadOrCreateDocument(part);
            var root = document.Root ?? throw new OfficeWorkerException("invalid_file", "metadata.xml 缺少根节点");
            var typeIndex = EnsureMetadataType(root);
            var valueIndex = EnsureFutureMetadata(root);
            var cellMetadataIndex = EnsureCellMetadata(root, typeIndex, valueIndex);
            SaveDocument(part, document);
            return cellMetadataIndex;
        }
    }

    public bool IsDynamicArrayMetadata(WorkbookPart workbookPart, int cellMetadataIndex)
    {
        var part = workbookPart.CellMetadataPart;
        if (part is null || cellMetadataIndex < 1) return false;
        var root = ReadOrCreateDocument(part).Root;
        if (root is null) return false;

        var metadataTypes = root.Element(Name("metadataTypes"));
        var cellMetadata = root.Element(Name("cellMetadata"));
        var block = cellMetadata?.Elements(Name("bk")).ElementAtOrDefault(cellMetadataIndex - 1);
        var record = block?.Element(Name("rc"));
        if (!int.TryParse((string?)record?.Attribute("t"), out var typeIndex) ||
            !int.TryParse((string?)record?.Attribute("v"), out var valueIndex)) return false;
        var type = metadataTypes?.Elements(Name("metadataType")).ElementAtOrDefault(typeIndex - 1);
        if (!string.Equals((string?)type?.Attribute("name"), "XLDAPR", StringComparison.Ordinal)) return false;
        var future = root.Elements(Name("futureMetadata"))
            .FirstOrDefault(element => string.Equals((string?)element.Attribute("name"), "XLDAPR", StringComparison.Ordinal));
        return future?.Elements(Name("bk")).ElementAtOrDefault(valueIndex)?
            .Descendants(XName.Get("dynamicArrayProperties", DynamicArrayNamespace))
            .Any(element => (string?)element.Attribute("fDynamic") == "1") == true;
    }

    private static CellMetadataPart EnsureMetadataPart(WorkbookPart workbookPart) =>
        workbookPart.CellMetadataPart ?? workbookPart.AddNewPart<CellMetadataPart>();

    private static XDocument ReadOrCreateDocument(CellMetadataPart part)
    {
        using var stream = part.GetStream(FileMode.OpenOrCreate, FileAccess.Read);
        if (stream.Length > 0) return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
        return new XDocument(
            new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(Name("metadata")));
    }

    private static void SaveDocument(CellMetadataPart part, XDocument document)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        document.Save(stream, SaveOptions.DisableFormatting);
    }

    private static int EnsureMetadataType(XElement root)
    {
        var types = root.Element(Name("metadataTypes"));
        if (types is null)
        {
            types = new XElement(Name("metadataTypes"));
            root.AddFirst(types);
        }

        var entries = types.Elements(Name("metadataType")).ToList();
        var index = entries.FindIndex(element => string.Equals((string?)element.Attribute("name"), "XLDAPR", StringComparison.Ordinal));
        if (index < 0)
        {
            types.Add(new XElement(Name("metadataType"),
                new XAttribute("name", "XLDAPR"),
                new XAttribute("minSupportedVersion", "120000"),
                new XAttribute("copy", "1"),
                new XAttribute("pasteAll", "1"),
                new XAttribute("pasteValues", "1"),
                new XAttribute("merge", "1"),
                new XAttribute("splitFirst", "1"),
                new XAttribute("rowColShift", "1"),
                new XAttribute("clearAll", "1"),
                new XAttribute("clearFormats", "1"),
                new XAttribute("clearContents", "1"),
                new XAttribute("clearComments", "1"),
                new XAttribute("assign", "1"),
                new XAttribute("coerce", "1"),
                new XAttribute("cellMeta", "1")));
            index = entries.Count;
        }
        types.SetAttributeValue("count", types.Elements(Name("metadataType")).Count());
        return index + 1;
    }

    private static int EnsureFutureMetadata(XElement root)
    {
        var future = root.Elements(Name("futureMetadata"))
            .FirstOrDefault(element => string.Equals((string?)element.Attribute("name"), "XLDAPR", StringComparison.Ordinal));
        if (future is null)
        {
            future = new XElement(Name("futureMetadata"), new XAttribute("name", "XLDAPR"));
            var cellMetadata = root.Element(Name("cellMetadata"));
            if (cellMetadata is null) root.Add(future); else cellMetadata.AddBeforeSelf(future);
        }

        var blocks = future.Elements(Name("bk")).ToList();
        var index = blocks.FindIndex(block => block
            .Descendants(XName.Get("dynamicArrayProperties", DynamicArrayNamespace))
            .Any(element => (string?)element.Attribute("fDynamic") == "1"));
        if (index < 0)
        {
            future.Add(new XElement(Name("bk"),
                new XElement(Name("extLst"),
                    new XElement(Name("ext"),
                        new XAttribute("uri", DynamicArrayExtensionUri),
                        new XElement(XName.Get("dynamicArrayProperties", DynamicArrayNamespace),
                            new XAttribute(XNamespace.Xmlns + "xda", DynamicArrayNamespace),
                            new XAttribute("fDynamic", "1"),
                            new XAttribute("fCollapsed", "0"))))));
            index = blocks.Count;
        }
        future.SetAttributeValue("count", future.Elements(Name("bk")).Count());
        return index;
    }

    private static int EnsureCellMetadata(XElement root, int typeIndex, int valueIndex)
    {
        var cellMetadata = root.Element(Name("cellMetadata"));
        if (cellMetadata is null)
        {
            cellMetadata = new XElement(Name("cellMetadata"));
            var valueMetadata = root.Element(Name("valueMetadata"));
            if (valueMetadata is null) root.Add(cellMetadata); else valueMetadata.AddBeforeSelf(cellMetadata);
        }

        var blocks = cellMetadata.Elements(Name("bk")).ToList();
        var index = blocks.FindIndex(block => block.Elements(Name("rc")).Any(record =>
            (string?)record.Attribute("t") == typeIndex.ToString() &&
            (string?)record.Attribute("v") == valueIndex.ToString()));
        if (index < 0)
        {
            cellMetadata.Add(new XElement(Name("bk"),
                new XElement(Name("rc"),
                    new XAttribute("t", typeIndex),
                    new XAttribute("v", valueIndex))));
            index = blocks.Count;
        }
        cellMetadata.SetAttributeValue("count", cellMetadata.Elements(Name("bk")).Count());
        return index + 1;
    }

    private static XName Name(string localName) => XName.Get(localName, MainNamespace);
}
