using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using Wengge.OfficeWorker.Office;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.OpenXml;

internal sealed class OpenXmlPresentationActionService
{
    private static readonly HashSet<string> AddOperations =
    ["addSlide", "addSlides", "appendSlide", "appendSlides", "addSlideContent"];
    private static readonly HashSet<string> Operations =
    ["createPresentation", "applyTheme", "deleteSlides", "addSlide", "addSlides", "appendSlide", "appendSlides", "addSlideContent", "normalizeLayouts", "alignShapes", "insertChart", "replacePictureSlot"];

    public static bool Supports(string operation) => Operations.Contains(operation);

    public object Execute(OfficeActionRequest request)
    {
        if (request.Operation == "createPresentation") return CreatePresentation(request);
        if (request.Operation == "applyTheme") return ApplyTheme(request);
        if (request.Operation == "deleteSlides") return DeleteSlides(request);
        if (AddOperations.Contains(request.Operation)) return AddSlides(request);
        if (request.Operation is "normalizeLayouts" or "alignShapes" or "insertChart" or "replacePictureSlot")
            return OfficeActionResults.NeedsCom(request, $"{request.Operation} 需要完整坐标、图表或媒体关系维护，需要 COM 执行");
        throw new OfficeWorkerException("unsupported_operation", $"不支持的 PowerPoint Open XML 操作: {request.Operation}");
    }

    private static object CreatePresentation(OfficeActionRequest request)
    {
        var output = RequiredOutput(request);
        Directory.CreateDirectory(Path.GetDirectoryName(output) ?? Environment.CurrentDirectory);
        if (File.Exists(output)) File.Delete(output);
        using var document = PresentationDocument.Create(output, DocumentFormat.OpenXml.PresentationDocumentType.Presentation);
        var presentationPart = document.AddPresentationPart();
        presentationPart.Presentation = new P.Presentation();
        var masterPart = presentationPart.AddNewPart<SlideMasterPart>();
        var layoutPart = masterPart.AddNewPart<SlideLayoutPart>();
        var themePart = masterPart.AddNewPart<ThemePart>();
        themePart.Theme = CreateTheme();
        layoutPart.SlideLayout = new P.SlideLayout(CreateCommonSlideData()) { Type = P.SlideLayoutValues.Blank, Preserve = true };
        layoutPart.AddPart(masterPart);
        var layoutRelationshipId = masterPart.GetIdOfPart(layoutPart);
        masterPart.SlideMaster = new P.SlideMaster(
            CreateCommonSlideData(),
            new P.ColorMap
            {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink,
            },
            new P.SlideLayoutIdList(new P.SlideLayoutId { Id = 2147483649U, RelationshipId = layoutRelationshipId }),
            new P.TextStyles(new P.TitleStyle(), new P.BodyStyle(), new P.OtherStyle()));
        var masterRelationshipId = presentationPart.GetIdOfPart(masterPart);
        presentationPart.Presentation.Append(
            new P.SlideMasterIdList(new P.SlideMasterId { Id = 2147483648U, RelationshipId = masterRelationshipId }),
            new P.SlideIdList(),
            new P.SlideSize { Cx = 12192000, Cy = 6858000, Type = P.SlideSizeValues.Screen16x9 },
            new P.NotesSize { Cx = 6858000, Cy = 9144000 },
            new P.DefaultTextStyle());
        var title = request.StringParam("title");
        var subtitle = request.StringParam("subtitle");
        AppendSlide(presentationPart, layoutPart, new SlideInput(title, subtitle), 256U);
        layoutPart.SlideLayout.Save();
        masterPart.SlideMaster.Save();
        themePart.Theme.Save();
        presentationPart.Presentation.Save();
        return Done(request, "已使用 .NET Open XML 创建 PowerPoint 演示文稿", output,
            ["ppt/presentation.xml", "ppt/slideMasters/slideMaster1.xml", "ppt/slideLayouts/slideLayout1.xml", "ppt/theme/theme1.xml", "ppt/slides/slide1.xml"],
            new { title, subtitle, slideCount = 1 });
    }

    private static object ApplyTheme(OfficeActionRequest request)
    {
        var output = PrepareCopy(request, "advanced", defaultToSource: false);
        using var document = PresentationDocument.Open(output, true);
        var presentation = document.PresentationPart ?? throw new OfficeWorkerException("invalid_file", "PowerPoint 文件缺少 presentation 部件");
        var color = NormalizeColor(request.StringParam("accentColor", "1F4E79"));
        var changed = new List<string>();
        foreach (var slidePart in presentation.SlideParts)
        {
            var slideChanged = false;
            foreach (var run in slidePart.Slide.Descendants<A.Run>())
            {
                run.RunProperties ??= new A.RunProperties();
                run.RunProperties.RemoveAllChildren<A.SolidFill>();
                run.RunProperties.PrependChild(new A.SolidFill(new A.RgbColorModelHex { Val = color }));
                slideChanged = true;
            }
            if (!slideChanged) continue;
            slidePart.Slide.Save();
            changed.Add(slidePart.Uri.ToString().TrimStart('/'));
        }
        return Done(request, "已应用 PowerPoint 主题色", output, changed, new { accentColor = color, slideCount = changed.Count });
    }

    private static object DeleteSlides(OfficeActionRequest request)
    {
        var output = PrepareCopy(request, "advanced", defaultToSource: false);
        using var document = PresentationDocument.Open(output, true);
        var presentationPart = document.PresentationPart ?? throw new OfficeWorkerException("invalid_file", "PowerPoint 文件缺少 presentation 部件");
        var list = presentationPart.Presentation.SlideIdList ?? throw new OfficeWorkerException("invalid_file", "PowerPoint 文件缺少幻灯片列表");
        var slideIds = list.Elements<P.SlideId>().ToArray();
        var indexes = DeleteIndexes(request, slideIds.Length);
        if (indexes.Count >= slideIds.Length) throw new OfficeWorkerException("invalid_params", "deleteSlides 至少需要保留一张幻灯片");
        var changed = new List<string>();
        foreach (var index in indexes.OrderByDescending(value => value))
        {
            var slideId = slideIds[index - 1];
            if (slideId.RelationshipId?.Value is { Length: > 0 } relationshipId)
            {
                var part = presentationPart.GetPartById(relationshipId);
                changed.Add(part.Uri.ToString().TrimStart('/'));
                presentationPart.DeletePart(part);
            }
            slideId.Remove();
        }
        presentationPart.Presentation.Save();
        return Done(request, $"已删除 {changed.Count} 张幻灯片", output, changed, new { deletedSlides = indexes });
    }

    private static object AddSlides(OfficeActionRequest request)
    {
        var slides = SlideInputs(request.Params);
        if (slides.Count == 0) throw new OfficeWorkerException("invalid_params", "addSlide 需要 params.title/body 或 params.slides");
        var output = PrepareCopy(request, "advanced", defaultToSource: true);
        using var document = PresentationDocument.Open(output, true);
        var presentationPart = document.PresentationPart ?? throw new OfficeWorkerException("invalid_file", "PowerPoint 文件缺少 presentation 部件");
        presentationPart.Presentation.SlideIdList ??= new P.SlideIdList();
        var layout = presentationPart.SlideParts.Select(part => part.SlideLayoutPart).FirstOrDefault(part => part is not null)
            ?? presentationPart.SlideMasterParts.SelectMany(part => part.SlideLayoutParts).FirstOrDefault();
        var nextId = presentationPart.Presentation.SlideIdList.Elements<P.SlideId>().Select(item => item.Id?.Value ?? 255U).DefaultIfEmpty(255U).Max() + 1;
        var changed = new List<string> { "ppt/presentation.xml" };
        foreach (var slide in slides)
        {
            var part = AppendSlide(presentationPart, layout, slide, nextId++);
            changed.Add(part.Uri.ToString().TrimStart('/'));
        }
        presentationPart.Presentation.Save();
        return Done(request, $"已添加 {slides.Count} 张幻灯片", output, changed, new { addedSlides = slides.Count });
    }

    private static SlidePart AppendSlide(PresentationPart presentationPart, SlideLayoutPart? layout, SlideInput input, uint id)
    {
        var slidePart = presentationPart.AddNewPart<SlidePart>();
        slidePart.Slide = CreateSlide(input.Title, input.Body);
        if (layout is not null) slidePart.AddPart(layout);
        var relationshipId = presentationPart.GetIdOfPart(slidePart);
        presentationPart.Presentation.SlideIdList ??= new P.SlideIdList();
        presentationPart.Presentation.SlideIdList.Append(new P.SlideId { Id = id, RelationshipId = relationshipId });
        slidePart.Slide.Save();
        return slidePart;
    }

    private static P.Slide CreateSlide(string title, string body)
    {
        var tree = CreateShapeTree();
        uint id = 2;
        if (title.Length > 0) tree.Append(CreateTextShape(id++, "Title", title, 685800, 457200, 10820400, 1143000, 2800, true));
        if (body.Length > 0) tree.Append(CreateTextShape(id, "Body", body, 914400, 2057400, 10287000, 3657600, 2000, false));
        return new P.Slide(new P.CommonSlideData(tree), new P.ColorMapOverride(new A.MasterColorMapping()));
    }

    private static P.Shape CreateTextShape(uint id, string name, string text, long x, long y, long cx, long cy, int fontSize, bool bold)
    {
        var paragraphs = text.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n').Select(line =>
            new A.Paragraph(
                new A.Run(new A.RunProperties { Language = "zh-CN", FontSize = fontSize, Bold = bold }, new A.Text(line)),
                new A.EndParagraphRunProperties { Language = "zh-CN", FontSize = fontSize })).ToArray();
        var textBody = new P.TextBody(
            new A.BodyProperties { Wrap = A.TextWrappingValues.Square, RightToLeftColumns = false },
            new A.ListStyle());
        textBody.Append(paragraphs);
        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.ShapeProperties(
                new A.Transform2D(new A.Offset { X = x, Y = y }, new A.Extents { Cx = cx, Cy = cy }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
                new A.NoFill()),
            textBody);
    }

    private static P.CommonSlideData CreateCommonSlideData() => new(CreateShapeTree()) { Name = string.Empty };

    private static P.ShapeTree CreateShapeTree() => new(
        new P.NonVisualGroupShapeProperties(
            new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
            new P.NonVisualGroupShapeDrawingProperties(),
            new P.ApplicationNonVisualDrawingProperties()),
        new P.GroupShapeProperties(new A.TransformGroup(
            new A.Offset { X = 0L, Y = 0L },
            new A.Extents { Cx = 0L, Cy = 0L },
            new A.ChildOffset { X = 0L, Y = 0L },
            new A.ChildExtents { Cx = 0L, Cy = 0L })));

    private static A.Theme CreateTheme() => new(
        new A.ThemeElements(
            new A.ColorScheme(
                new A.Dark1Color(new A.SystemColor { Val = A.SystemColorValues.WindowText, LastColor = "000000" }),
                new A.Light1Color(new A.SystemColor { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" }),
                new A.Dark2Color(new A.RgbColorModelHex { Val = "1F1F1F" }),
                new A.Light2Color(new A.RgbColorModelHex { Val = "F2F2F2" }),
                new A.Accent1Color(new A.RgbColorModelHex { Val = "1F4E79" }),
                new A.Accent2Color(new A.RgbColorModelHex { Val = "70AD47" }),
                new A.Accent3Color(new A.RgbColorModelHex { Val = "ED7D31" }),
                new A.Accent4Color(new A.RgbColorModelHex { Val = "5B9BD5" }),
                new A.Accent5Color(new A.RgbColorModelHex { Val = "A5A5A5" }),
                new A.Accent6Color(new A.RgbColorModelHex { Val = "FFC000" }),
                new A.Hyperlink(new A.RgbColorModelHex { Val = "0563C1" }),
                new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = "954F72" })) { Name = "Wengge" },
            new A.FontScheme(
                new A.MajorFont(new A.LatinFont { Typeface = "Aptos Display" }, new A.EastAsianFont { Typeface = "微软雅黑" }, new A.ComplexScriptFont { Typeface = "Arial" }),
                new A.MinorFont(new A.LatinFont { Typeface = "Aptos" }, new A.EastAsianFont { Typeface = "微软雅黑" }, new A.ComplexScriptFont { Typeface = "Arial" })) { Name = "Wengge" },
            new A.FormatScheme(
                new A.FillStyleList(ThemeFill(), ThemeFill(), ThemeFill()),
                new A.LineStyleList(ThemeLine(9525), ThemeLine(25400), ThemeLine(38100)),
                new A.EffectStyleList(ThemeEffect(), ThemeEffect(), ThemeEffect()),
                new A.BackgroundFillStyleList(ThemeFill(), ThemeFill(), ThemeFill())) { Name = "Wengge" }))
    { Name = "Wengge Theme" };

    private static A.SolidFill ThemeFill() => new(new A.SchemeColor { Val = A.SchemeColorValues.PhColor });
    private static A.Outline ThemeLine(int width) => new(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })) { Width = width };
    private static A.EffectStyle ThemeEffect() => new(new A.EffectList());

    private static List<SlideInput> SlideInputs(JsonElement parameters)
    {
        var result = new List<SlideInput>();
        if (parameters.ValueKind != JsonValueKind.Object) return result;
        if (parameters.TryGetProperty("slides", out var slides) && slides.ValueKind == JsonValueKind.Array)
        {
            foreach (var slide in slides.EnumerateArray())
            {
                if (slide.ValueKind != JsonValueKind.Object) continue;
                var title = String(slide, "title");
                var body = String(slide, "body");
                if (body.Length == 0 && slide.TryGetProperty("bullets", out var bullets) && bullets.ValueKind == JsonValueKind.Array)
                    body = string.Join('\n', bullets.EnumerateArray().Select(item => item.ValueKind == JsonValueKind.String ? item.GetString() : item.ToString()).Where(item => !string.IsNullOrWhiteSpace(item)).Select(item => $"• {item}"));
                if (title.Length > 0 || body.Length > 0) result.Add(new SlideInput(title, body));
            }
        }
        if (result.Count == 0)
        {
            var title = String(parameters, "title");
            var body = String(parameters, "body");
            if (body.Length == 0) body = String(parameters, "text");
            if (title.Length > 0 || body.Length > 0) result.Add(new SlideInput(title, body));
        }
        return result;
    }

    private static List<int> DeleteIndexes(OfficeActionRequest request, int slideCount)
    {
        var indexes = new List<int>();
        var slides = request.Param("slides");
        if (slides.ValueKind == JsonValueKind.Array)
            indexes.AddRange(slides.EnumerateArray().Select(item => item.TryGetInt32(out var value) ? value : 0));
        if (indexes.Count == 0)
        {
            var from = request.IntParam("from", request.IntParam("start"));
            var to = request.IntParam("to", request.IntParam("end", from));
            if (from > 0) indexes.AddRange(Enumerable.Range(Math.Min(from, to), Math.Abs(to - from) + 1));
        }
        if (indexes.Count == 0 && request.Target is { Length: > 0 } target)
        {
            var value = target.StartsWith("slides:", StringComparison.OrdinalIgnoreCase) ? target[7..]
                : target.StartsWith("slide:", StringComparison.OrdinalIgnoreCase) ? target[6..] : string.Empty;
            var parts = value.Split('-', 2);
            if (int.TryParse(parts[0], out var from))
            {
                var to = parts.Length > 1 && int.TryParse(parts[1], out var parsed) ? parsed : from;
                indexes.AddRange(Enumerable.Range(Math.Min(from, to), Math.Abs(to - from) + 1));
            }
        }
        indexes = indexes.Where(index => index > 0).Distinct().Order().ToList();
        if (indexes.Count == 0) throw new OfficeWorkerException("invalid_params", "deleteSlides 需要 params.slides、params.from/to 或 target: slide:2-6");
        var overflow = indexes.FirstOrDefault(index => index > slideCount);
        if (overflow > 0) throw new OfficeWorkerException("invalid_params", $"幻灯片序号超出范围: {overflow}");
        return indexes;
    }

    private static string PrepareCopy(OfficeActionRequest request, string suffix, bool defaultToSource)
    {
        var source = RequireFile(request.FilePath);
        var output = !string.IsNullOrWhiteSpace(request.OutputPath) ? Path.GetFullPath(request.OutputPath)
            : defaultToSource ? source : Path.Combine(Path.GetDirectoryName(source) ?? Environment.CurrentDirectory, $"{Path.GetFileNameWithoutExtension(source)}-{suffix}{Path.GetExtension(source)}");
        Directory.CreateDirectory(Path.GetDirectoryName(output) ?? Environment.CurrentDirectory);
        if (!string.Equals(source, output, StringComparison.OrdinalIgnoreCase)) File.Copy(source, output, true);
        return output;
    }

    private static string RequiredOutput(OfficeActionRequest request)
    {
        var path = request.OutputPath ?? request.FilePath;
        return string.IsNullOrWhiteSpace(path) ? throw new OfficeWorkerException("invalid_params", "createPresentation 需要 filePath 或 outputPath") : Path.GetFullPath(path);
    }

    private static string RequireFile(string? path)
    {
        if (string.IsNullOrWhiteSpace(path)) throw new OfficeWorkerException("invalid_params", "缺少 filePath");
        var fullPath = Path.GetFullPath(path);
        return File.Exists(fullPath) ? fullPath : throw new OfficeWorkerException("file_not_found", $"Office 文件不存在: {fullPath}");
    }

    private static object Done(OfficeActionRequest request, string summary, string output, IEnumerable<string> parts, object? data = null) =>
        OfficeActionResults.Done(request, "openxml", summary, data ?? new { outputPath = output, changedParts = parts.ToArray() },
            parts.Select(part => new OfficeChange("openxml-part", part, $"已更新 {part}")), output);

    private static string String(JsonElement value, string property) => value.TryGetProperty(property, out var item) && item.ValueKind == JsonValueKind.String ? item.GetString() ?? string.Empty : string.Empty;
    private static string NormalizeColor(string value) { var color = value.Trim().TrimStart('#'); return color.Length == 6 && color.All(Uri.IsHexDigit) ? color.ToUpperInvariant() : "1F4E79"; }
    private sealed record SlideInput(string Title, string Body);
}
