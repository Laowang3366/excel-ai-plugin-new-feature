using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.Json;
using Wengge.OfficeWorker.Protocol;
using Wengge.OfficeWorker.Runtime;

namespace Wengge.OfficeWorker.Tests;

public sealed class OpenXmlWorkerTests : IDisposable
{
    private readonly string root = Path.Combine(Path.GetTempPath(), $"wengge-openxml-{Guid.NewGuid():N}");

    public OpenXmlWorkerTests() => Directory.CreateDirectory(root);

    [Fact]
    public async Task InspectsAndReplacesWordText()
    {
        var source = Path.Combine(root, "source.docx");
        var output = Path.Combine(root, "output.docx");
        CreateDocument(source, "hello Office");
        using var worker = OfficeWorkerHost.Create();

        var inspected = await worker.DispatchAsync(Request("openxml.inspect", new { filePath = source }), CancellationToken.None);
        var inspectJson = JsonSerializer.SerializeToElement(inspected, JsonOptions);
        Assert.Contains("hello Office", inspectJson.GetProperty("textPreview").GetString());

        var replaced = await worker.DispatchAsync(Request("openxml.replaceText", new
        {
            filePath = source,
            outputPath = output,
            findText = "Office",
            replaceText = "Worker",
        }), CancellationToken.None);
        var replaceJson = JsonSerializer.SerializeToElement(replaced, JsonOptions);
        Assert.Equal(1, replaceJson.GetProperty("replacements").GetInt32());
        Assert.Equal("hello Worker", ReadDocument(output));
    }

    [Fact]
    public async Task ReplacesWordBodyHeaderAndFooterInPlace()
    {
        var path = Path.Combine(root, "in-place.docx");
        CreateDocument(path, "Office body", "Office header", "Office footer");
        using var worker = OfficeWorkerHost.Create();

        var inspected = await worker.DispatchAsync(Request("openxml.inspect", new { filePath = path }), CancellationToken.None);
        var inspectJson = JsonSerializer.SerializeToElement(inspected, JsonOptions);
        var partNames = inspectJson.GetProperty("textParts").EnumerateArray()
            .Select(part => part.GetProperty("partName").GetString()).ToArray();
        Assert.Contains("word/header1.xml", partNames);
        Assert.Contains("word/footer1.xml", partNames);

        var replaced = await worker.DispatchAsync(Request("openxml.replaceText", new
        {
            filePath = path,
            outputPath = path,
            findText = "Office",
            replaceText = "Worker",
        }), CancellationToken.None);
        var replaceJson = JsonSerializer.SerializeToElement(replaced, JsonOptions);
        Assert.Equal(3, replaceJson.GetProperty("replacements").GetInt32());

        using var document = WordprocessingDocument.Open(path, false);
        Assert.Equal("Worker body", document.MainDocumentPart!.Document.InnerText);
        Assert.Equal("Worker header", document.MainDocumentPart.HeaderParts.Single().Header.InnerText);
        Assert.Equal("Worker footer", document.MainDocumentPart.FooterParts.Single().Footer.InnerText);
    }

    public void Dispose()
    {
        try { Directory.Delete(root, recursive: true); } catch { }
    }

    private static void CreateDocument(string path, string text, string? headerText = null, string? footerText = null)
    {
        using var document = WordprocessingDocument.Create(path, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        var section = new SectionProperties();
        if (headerText is not null)
        {
            var header = main.AddNewPart<HeaderPart>();
            header.Header = new Header(new Paragraph(new Run(new Text(headerText))));
            header.Header.Save();
            section.Append(new HeaderReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(header) });
        }
        if (footerText is not null)
        {
            var footer = main.AddNewPart<FooterPart>();
            footer.Footer = new Footer(new Paragraph(new Run(new Text(footerText))));
            footer.Footer.Save();
            section.Append(new FooterReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(footer) });
        }
        main.Document = new Document(new Body(new Paragraph(new Run(new Text(text))), section));
        main.Document.Save();
    }

    private static string ReadDocument(string path)
    {
        using var document = WordprocessingDocument.Open(path, false);
        return string.Concat(document.MainDocumentPart!.Document.Descendants<Text>().Select(text => text.Text));
    }

    private static RpcRequest Request(string method, object parameters) =>
        new("test", method, JsonSerializer.SerializeToElement(parameters, JsonOptions));

    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);
}
