using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using S = DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using Wengge.OfficeWorker.Protocol;
using Wengge.OfficeWorker.Runtime;
using Wengge.OfficeWorker.Office;

namespace Wengge.OfficeWorker.Tests;

public sealed class AdvancedOpenXmlActionTests : IDisposable
{
    private readonly string root = Path.Combine(Path.GetTempPath(), $"wengge-openxml-actions-{Guid.NewGuid():N}");
    public AdvancedOpenXmlActionTests() => Directory.CreateDirectory(root);

    [Fact]
    public void NormalizesTrackedBookmarkBoundariesInsideInsertedRun()
    {
        var path = Path.Combine(root, "tracked-bookmark.docx");
        using (var package = WordprocessingDocument.Create(path, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = package.AddMainDocumentPart();
            main.Document = new W.Document(new W.Body(new W.Paragraph(
                new W.BookmarkStart { Id = "0", Name = "ManualKeep" },
                new W.InsertedRun(
                    new W.Run(new W.Text("人工保留段落")))
                {
                    Id = "1",
                    Author = "test",
                    Date = DateTime.UtcNow,
                },
                new W.BookmarkEnd { Id = "0" })));
            main.Document.Save();
        }

        var normalized = WordTrackedBookmarkPackageUpdater.Normalize(path,
            [new WordTrackedBookmarkUpdate("ManualKeep", "人工保留段落")]);

        Assert.Equal(["ManualKeep"], normalized);
        using var document = WordprocessingDocument.Open(path, false);
        var paragraph = document.MainDocumentPart!.Document.Descendants<W.Paragraph>().Single();
        var start = paragraph.Elements<W.BookmarkStart>().Single();
        var end = paragraph.Elements<W.BookmarkEnd>().Single();
        var insertion = paragraph.Elements<W.InsertedRun>().Single();
        var anchor = paragraph.Elements<W.Run>().Single();
        Assert.Equal("\u200B", anchor.InnerText);
        Assert.True(paragraph.ChildElements.ToList().IndexOf(start) < paragraph.ChildElements.ToList().IndexOf(anchor));
        Assert.True(paragraph.ChildElements.ToList().IndexOf(end) > paragraph.ChildElements.ToList().IndexOf(insertion));
        var bookmarks = WordTrackedBookmarkPackageUpdater.Read(path);
        Assert.Contains(bookmarks, item => item.Name == "ManualKeep" && item.Text == "人工保留段落");
        AssertValid(document);
    }

    [Fact]
    public async Task CreatesAndEditsSpreadsheetThroughUnifiedRoute()
    {
        var path = Path.Combine(root, "book.xlsx");
        using var worker = OfficeWorkerHost.Create();
        await Execute(worker, new
        {
            app = "excel", action = "insert", operation = "createWorkbook", filePath = path,
            target = "range:Data!A1:C3",
            @params = new { sheetNames = new[] { "Data" }, values = new object[][] { ["Name", "Score"], ["Alice", 95], ["Dynamic", "=FILTER(A:A,A:A<>\"\")"] } },
        });
        await Execute(worker, new
        {
            app = "excel", action = "edit", operation = "setDataValidation", filePath = path, outputPath = path,
            target = "range:Data!C2:C10", @params = new { values = new[] { "通过", "失败" } },
        });
        await Execute(worker, new
        {
            app = "excel", action = "style", operation = "applyConditionalFormatting", filePath = path, outputPath = path,
            target = "range:Data!B2:B10", @params = new { fillColor = "FFF2CC" },
        });

        using var document = SpreadsheetDocument.Open(path, false);
        var worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet;
        Assert.Equal("Alice", worksheet.Descendants<S.Cell>().Single(cell => cell.CellReference == "A2").InlineString!.InnerText);
        Assert.Contains(worksheet.Descendants<S.DataValidation>(), item => item.SequenceOfReferences?.InnerText == "C2:C10");
        Assert.Contains(worksheet.Descendants<S.ConditionalFormatting>(), item => item.SequenceOfReferences?.InnerText == "B2:B10");
        Assert.Contains(worksheet.Descendants<S.CellFormula>(), formula => formula.Text.Contains("_xlfn._xlws.FILTER", StringComparison.Ordinal));
        AssertValid(document);
    }

    [Fact]
    public async Task CreatesStylesAndAddsHeaderFooterToWordDocument()
    {
        var path = Path.Combine(root, "report.docx");
        using var worker = OfficeWorkerHost.Create();
        await Execute(worker, new
        {
            app = "word", action = "insert", operation = "createDocument", filePath = path,
            @params = new { title = "测试报告", paragraphs = new[] { "一、概览", "正文" } },
        });
        await Execute(worker, new
        {
            app = "word", action = "style", operation = "applyHeadingStyles", filePath = path, outputPath = path,
            @params = new { startsWith = "一、", level = 1 },
        });
        await Execute(worker, new
        {
            app = "word", action = "edit", operation = "setHeaderFooter", filePath = path, outputPath = path,
            @params = new { kind = "header", text = "页眉" },
        });
        await Execute(worker, new
        {
            app = "word", action = "edit", operation = "setHeaderFooter", filePath = path, outputPath = path,
            @params = new { kind = "footer", text = "页脚" },
        });

        using var document = WordprocessingDocument.Open(path, false);
        var main = document.MainDocumentPart!;
        Assert.Contains(main.Document.Descendants<W.ParagraphStyleId>(), style => style.Val == "Heading1");
        Assert.Equal("页眉", main.HeaderParts.Single().Header.InnerText);
        Assert.Equal("页脚", main.FooterParts.Single().Footer.InnerText);
        var sectionChildren = main.Document.Body!.Elements<W.SectionProperties>().Single().ChildElements.ToArray();
        Assert.True(Array.FindIndex(sectionChildren, item => item is W.HeaderReference) < Array.FindIndex(sectionChildren, item => item is W.FooterReference));
        AssertValid(document);
    }

    [Fact]
    public async Task CreatesAddsThemesAndDeletesPresentationSlides()
    {
        var path = Path.Combine(root, "slides.pptx");
        using var worker = OfficeWorkerHost.Create();
        await Execute(worker, new
        {
            app = "presentation", action = "insert", operation = "createPresentation", filePath = path,
            @params = new { title = "健康饮食", subtitle = "均衡膳食" },
        });
        await Execute(worker, new
        {
            app = "presentation", action = "insert", operation = "addSlides", filePath = path,
            @params = new { slides = new object[] { new { title = "营养原则", bullets = new[] { "均衡膳食", "适量运动" } }, new { title = "每日建议", body = "早餐\n午餐\n晚餐" } } },
        });
        await Execute(worker, new
        {
            app = "presentation", action = "style", operation = "applyTheme", filePath = path, outputPath = path,
            @params = new { accentColor = "1F4E79" },
        });
        await Execute(worker, new
        {
            app = "presentation", action = "edit", operation = "deleteSlides", filePath = path, outputPath = path,
            @params = new { slides = new[] { 2 } },
        });

        using var document = PresentationDocument.Open(path, false);
        var presentation = document.PresentationPart!;
        Assert.Equal(2, presentation.Presentation.SlideIdList!.Elements<P.SlideId>().Count());
        Assert.Contains(presentation.SlideParts.SelectMany(part => part.Slide.Descendants<A.Text>()), text => text.Text == "每日建议");
        Assert.All(presentation.SlideParts.SelectMany(part => part.Slide.Descendants<A.RunProperties>()), properties =>
            Assert.Contains(properties.Elements<A.SolidFill>(), fill => fill.RgbColorModelHex?.Val == "1F4E79"));
        AssertValid(document);
    }

    [Fact]
    public async Task CreatesAndUpdatesSpeakerNotesWithoutAnExistingNotesTemplate()
    {
        var path = Path.Combine(root, "notes.pptx");
        using var worker = OfficeWorkerHost.Create();
        await Execute(worker, new
        {
            app = "presentation", action = "insert", operation = "createPresentation", filePath = path,
            @params = new { title = "第一页", subtitle = "无备注模板" },
        });
        await Execute(worker, new
        {
            app = "presentation", action = "insert", operation = "addSlides", filePath = path,
            @params = new { slides = new object[] { new { title = "第二页", body = "正文" } } },
        });

        Assert.True(PresentationNotesPackageUpdater.Apply(path,
        [
            new PresentationNoteUpdate(1, "第一页备注", false),
            new PresentationNoteUpdate(2, "第二页备注", false),
        ]));
        Assert.True(PresentationNotesPackageUpdater.Apply(path,
        [
            new PresentationNoteUpdate(2, "追加内容", true),
        ]));

        using var document = PresentationDocument.Open(path, false);
        var presentation = document.PresentationPart!;
        var slides = presentation.Presentation.SlideIdList!.Elements<P.SlideId>()
            .Select(id => presentation.GetPartById(id.RelationshipId!)).Cast<SlidePart>().ToArray();
        Assert.All(slides, slide => Assert.NotNull(slide.NotesSlidePart));
        Assert.Equal("第一页备注", NotesText(slides[0]));
        Assert.Contains("第二页备注", NotesText(slides[1]));
        Assert.Contains("追加内容", NotesText(slides[1]));
        var snapshots = PresentationNotesPackageUpdater.Read(path, true, 1);
        Assert.Equal(2, snapshots.Count);
        Assert.Equal("第一页备注", snapshots[0].NotesText);
        Assert.Contains("追加内容", snapshots[1].NotesText);
        Assert.NotNull(presentation.NotesMasterPart);
        Assert.All(slides, slide => Assert.Same(presentation.NotesMasterPart, slide.NotesSlidePart!.NotesMasterPart));
        AssertValid(document);
    }

    [Fact]
    public async Task RoutesExplicitWpsSpeakerNotesThroughPackageFallbackWithoutCom()
    {
        var path = Path.Combine(root, "wps-notes.pptx");
        using var worker = OfficeWorkerHost.Create();
        await Execute(worker, new
        {
            app = "presentation", action = "insert", operation = "createPresentation", filePath = path,
            @params = new { title = "第一页", subtitle = "WPS 包级备注" },
        });
        await Execute(worker, new
        {
            app = "presentation", action = "insert", operation = "addSlides", filePath = path,
            @params = new { slides = new object[] { new { title = "第二页", body = "正文" } } },
        });

        var write = await ExecuteOfficeAction(worker, new
        {
            app = "presentation", action = "edit", operation = "setSpeakerNotes", filePath = path,
            @params = new
            {
                host = "wps",
                notesBySlide = new object[]
                {
                    new { slideIndex = 1, text = "第一页备注" },
                    new { slideIndex = 2, text = "第二页备注" },
                },
            },
        });
        Assert.Equal("openxml", write.GetProperty("engine").GetString());
        Assert.True(write.GetProperty("data").GetProperty("notesPackageFallback").GetBoolean());

        var inspect = await ExecuteOfficeAction(worker, new
        {
            app = "presentation", action = "inspect", operation = "inspectSpeakerNotes", filePath = path,
            @params = new { host = "wps", allSlides = true },
        });
        Assert.Equal("openxml", inspect.GetProperty("engine").GetString());
        Assert.Equal(0, inspect.GetProperty("data").GetProperty("summary").GetProperty("missingNotes").GetInt32());
    }

    [Fact]
    public async Task ParsesSpreadsheetWordAndPresentationForLocalDocumentConsumers()
    {
        using var worker = OfficeWorkerHost.Create();
        var workbookPath = Path.Combine(root, "parse.xlsx");
        await Execute(worker, new
        {
            app = "excel", action = "insert", operation = "createWorkbook", filePath = workbookPath,
            target = "range:People!A1:B2",
            @params = new { sheetNames = new[] { "People" }, values = new object[][] { ["Name", "Age"], ["Ada", 36] } },
        });
        var workbook = await Parse(worker, workbookPath);
        var workbookChunk = workbook.GetProperty("chunks")[0];
        Assert.Equal("People", workbookChunk.GetProperty("metadata").GetProperty("sheetName").GetString());
        Assert.Equal("A1:B2", workbookChunk.GetProperty("metadata").GetProperty("tableRange").GetString());
        Assert.Contains("Ada | 36", workbookChunk.GetProperty("content").GetString());

        var documentPath = Path.Combine(root, "parse.docx");
        await Execute(worker, new
        {
            app = "word", action = "insert", operation = "createDocument", filePath = documentPath,
            @params = new { title = "区域汇总", paragraphs = new[] { "正文内容" } },
        });
        using (var document = WordprocessingDocument.Open(documentPath, true))
        {
            document.MainDocumentPart!.Document.Body!.InsertBefore(new W.Table(
                new W.TableRow(new W.TableCell(new W.Paragraph(new W.Run(new W.Text("字段"))))),
                new W.TableRow(new W.TableCell(new W.Paragraph(new W.Run(new W.Text("销售额")))))),
                document.MainDocumentPart.Document.Body.Elements<W.SectionProperties>().Single());
            document.MainDocumentPart.Document.Save();
        }
        var word = await Parse(worker, documentPath);
        var wordChunk = word.GetProperty("chunks")[0];
        Assert.Contains("区域汇总", wordChunk.GetProperty("content").GetString());
        Assert.Equal("字段", wordChunk.GetProperty("metadata").GetProperty("rows")[0][0].GetString());

        var presentationPath = Path.Combine(root, "parse.pptx");
        await Execute(worker, new
        {
            app = "presentation", action = "insert", operation = "createPresentation", filePath = presentationPath,
            @params = new { title = "知识库演示页", subtitle = "支持提取 PPT 文本" },
        });
        var presentation = await Parse(worker, presentationPath);
        var slideChunk = presentation.GetProperty("chunks")[0];
        Assert.Equal(1, slideChunk.GetProperty("metadata").GetProperty("slideNumber").GetInt32());
        Assert.Contains("知识库演示页", slideChunk.GetProperty("content").GetString());
    }

    private static async Task<JsonElement> Execute(OfficeWorkerHost worker, object parameters)
    {
        var result = await worker.DispatchAsync(new RpcRequest("test", "openxml.action.execute", JsonSerializer.SerializeToElement(parameters, JsonOptions)), CancellationToken.None);
        var json = JsonSerializer.SerializeToElement(result, JsonOptions);
        Assert.Equal("done", json.GetProperty("status").GetString());
        Assert.Equal("openxml", json.GetProperty("engine").GetString());
        return json;
    }

    private static async Task<JsonElement> Parse(OfficeWorkerHost worker, string filePath)
    {
        var result = await worker.DispatchAsync(new RpcRequest(
            "test",
            "openxml.parseDocument",
            JsonSerializer.SerializeToElement(new { filePath }, JsonOptions)), CancellationToken.None);
        return JsonSerializer.SerializeToElement(result, JsonOptions);
    }

    private static async Task<JsonElement> ExecuteOfficeAction(OfficeWorkerHost worker, object parameters)
    {
        var result = await worker.DispatchAsync(new RpcRequest("test", "office.action.execute", JsonSerializer.SerializeToElement(parameters, JsonOptions)), CancellationToken.None);
        var json = JsonSerializer.SerializeToElement(result, JsonOptions);
        Assert.Equal("done", json.GetProperty("status").GetString());
        return json;
    }

    private static void AssertValid(OpenXmlPackage document)
    {
        var errors = new OpenXmlValidator().Validate(document).Take(20).Select(error => $"{error.Path?.XPath}: {error.Description}").ToArray();
        Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors));
    }

    private static string NotesText(SlidePart slidePart)
    {
        var body = slidePart.NotesSlidePart!.NotesSlide.Descendants<P.Shape>().Single(shape =>
            shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == P.PlaceholderValues.Body);
        return string.Concat(body.Descendants<A.Text>().Select(text => text.Text));
    }

    public void Dispose()
    {
        try { Directory.Delete(root, true); } catch { }
    }

    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);
}
