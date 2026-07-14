using DocumentFormat.OpenXml.Packaging;
using Wengge.OfficeWorker.Protocol;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace Wengge.OfficeWorker.Office;

internal sealed record PresentationNoteUpdate(int SlideIndex, string Text, bool Append);
internal sealed record PresentationPackageNote(int SlideIndex, string SlideText, string NotesText);

internal static class PresentationNotesPackageUpdater
{
    public static IReadOnlyList<PresentationPackageNote> Read(string? filePath, bool allSlides, int slideIndex)
    {
        if (string.IsNullOrWhiteSpace(filePath)) throw new OfficeWorkerException("invalid_params", "备注包级读取缺少演示文稿路径");
        using var document = PresentationDocument.Open(Path.GetFullPath(filePath), false);
        var presentationPart = document.PresentationPart
            ?? throw new OfficeWorkerException("invalid_presentation", "演示文稿缺少 presentation 部件");
        var slideParts = OrderedSlideParts(presentationPart);
        var indexes = allSlides ? Enumerable.Range(1, slideParts.Count) : [slideIndex];
        return indexes.Select(index =>
        {
            if (index < 1 || index > slideParts.Count)
                throw new OfficeWorkerException("slide_not_found", $"备注页序号超出范围: {index}");
            var slidePart = slideParts[index - 1];
            var slideText = string.Join(" ", slidePart.Slide.Descendants<A.Text>().Select(node => node.Text).Where(text => !string.IsNullOrWhiteSpace(text)));
            var body = slidePart.NotesSlidePart?.NotesSlide?.CommonSlideData?.ShapeTree?.Elements<P.Shape>().FirstOrDefault(shape =>
                shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == P.PlaceholderValues.Body);
            var notesText = body is null ? string.Empty : string.Concat(body.Descendants<A.Text>().Select(node => node.Text));
            return new PresentationPackageNote(index, slideText, notesText);
        }).ToArray();
    }

    public static bool Apply(string? filePath, IReadOnlyList<PresentationNoteUpdate> updates)
    {
        if (updates.Count == 0) return false;
        if (string.IsNullOrWhiteSpace(filePath)) throw new OfficeWorkerException("invalid_params", "备注包级回写缺少演示文稿路径");
        var fullPath = System.IO.Path.GetFullPath(filePath);
        Exception? lastError = null;
        for (var attempt = 0; attempt < 10; attempt++)
        {
            try { ApplyCore(fullPath, updates); return true; }
            catch (IOException exception)
            {
                lastError = exception;
                Thread.Sleep(200);
            }
        }
        throw new OfficeWorkerException("file_locked", $"WPS 保存后演示文稿仍被占用: {fullPath}", null, lastError);
    }

    private static void ApplyCore(string filePath, IReadOnlyList<PresentationNoteUpdate> updates)
    {
        using var document = PresentationDocument.Open(filePath, true);
        var presentationPart = document.PresentationPart
            ?? throw new OfficeWorkerException("invalid_presentation", "演示文稿缺少 presentation 部件");
        var slideParts = OrderedSlideParts(presentationPart);
        var template = slideParts.Select(part => part.NotesSlidePart).FirstOrDefault(part => part?.NotesSlide is not null);
        var notesMaster = template?.NotesMasterPart ?? presentationPart.NotesMasterPart;
        foreach (var update in updates)
        {
            if (update.SlideIndex < 1 || update.SlideIndex > slideParts.Count)
                throw new OfficeWorkerException("slide_not_found", $"备注页序号超出范围: {update.SlideIndex}");
            var slidePart = slideParts[update.SlideIndex - 1];
            var notesPart = slidePart.NotesSlidePart;
            if (notesPart is null)
            {
                notesPart = CreateNotesPart(slidePart, template);
                notesMaster ??= CreateNotesMaster(presentationPart);
                if (notesPart.NotesMasterPart is null) notesPart.AddPart(notesMaster);
            }
            else if (notesPart.NotesMasterPart is null)
            {
                notesMaster ??= CreateNotesMaster(presentationPart);
                notesPart.AddPart(notesMaster);
            }
            EnsureBodyShape(notesPart);
            UpdateText(notesPart, update);
        }
        presentationPart.Presentation.Save();
    }

    private static List<SlidePart> OrderedSlideParts(PresentationPart presentationPart)
    {
        var slideIds = presentationPart.Presentation?.SlideIdList?.Elements<P.SlideId>()
            ?? throw new OfficeWorkerException("invalid_presentation", "演示文稿缺少幻灯片列表");
        return slideIds.Select(slideId => slideId.RelationshipId?.Value)
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .Select(id => presentationPart.GetPartById(id!))
            .OfType<SlidePart>()
            .ToList();
    }

    private static NotesSlidePart CreateNotesPart(SlidePart slidePart, NotesSlidePart? template)
    {
        var notesPart = slidePart.AddNewPart<NotesSlidePart>();
        notesPart.NotesSlide = template?.NotesSlide is not null
            ? (P.NotesSlide)template.NotesSlide.CloneNode(true)
            : CreateNotesSlide();
        notesPart.AddPart(slidePart);
        return notesPart;
    }

    private static P.NotesSlide CreateNotesSlide() => new(
        new P.CommonSlideData(
            new P.ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(
                    new A.TransformGroup(
                        new A.Offset { X = 0L, Y = 0L },
                        new A.Extents { Cx = 0L, Cy = 0L },
                        new A.ChildOffset { X = 0L, Y = 0L },
                        new A.ChildExtents { Cx = 0L, Cy = 0L })),
                CreateBodyShape())),
        new P.ColorMapOverride(new A.MasterColorMapping()));

    private static P.Shape CreateBodyShape() => new(
        new P.NonVisualShapeProperties(
            new P.NonVisualDrawingProperties { Id = 2U, Name = "Notes Placeholder" },
            new P.NonVisualShapeDrawingProperties(),
            new P.ApplicationNonVisualDrawingProperties(
                new P.PlaceholderShape { Type = P.PlaceholderValues.Body, Index = 1U })),
        new P.ShapeProperties(),
        new P.TextBody(
            new A.BodyProperties(),
            new A.ListStyle(),
            new A.Paragraph(new A.EndParagraphRunProperties { Language = "zh-CN" })));

    private static void EnsureBodyShape(NotesSlidePart notesPart)
    {
        var shapeTree = notesPart.NotesSlide?.CommonSlideData?.ShapeTree
            ?? throw new OfficeWorkerException("invalid_notes", "备注页缺少形状树");
        if (shapeTree.Elements<P.Shape>().Any(shape =>
            shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == P.PlaceholderValues.Body)) return;
        shapeTree.Append(CreateBodyShape());
    }

    private static NotesMasterPart CreateNotesMaster(PresentationPart presentationPart)
    {
        var notesMaster = presentationPart.AddNewPart<NotesMasterPart>();
        notesMaster.NotesMaster = new P.NotesMaster(
            new P.CommonSlideData(
                new P.ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 1U, Name = "Notes Master" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new P.ApplicationNonVisualDrawingProperties()),
                    new P.GroupShapeProperties())),
            new P.ColorMap
            {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink,
            });
        var existingTheme = presentationPart.SlideMasterParts.Select(part => part.ThemePart).FirstOrDefault(part => part is not null);
        if (existingTheme is not null) notesMaster.AddPart(existingTheme);
        else
        {
            var themePart = notesMaster.AddNewPart<ThemePart>();
            themePart.Theme = CreateFallbackTheme();
            themePart.Theme.Save();
        }
        notesMaster.NotesMaster.Save();

        var presentation = presentationPart.Presentation;
        presentation.NotesMasterIdList ??= new P.NotesMasterIdList();
        presentation.NotesMasterIdList.Append(new P.NotesMasterId { Id = presentationPart.GetIdOfPart(notesMaster) });
        return notesMaster;
    }

    private static A.Theme CreateFallbackTheme() => new(
        new A.ThemeElements(
            new A.ColorScheme(
                new A.Dark1Color(new A.SystemColor { Val = A.SystemColorValues.WindowText }),
                new A.Light1Color(new A.SystemColor { Val = A.SystemColorValues.Window }),
                new A.Dark2Color(new A.RgbColorModelHex { Val = "1F1F1F" }),
                new A.Light2Color(new A.RgbColorModelHex { Val = "E7E6E6" }),
                new A.Accent1Color(new A.RgbColorModelHex { Val = "4472C4" }),
                new A.Accent2Color(new A.RgbColorModelHex { Val = "ED7D31" }),
                new A.Accent3Color(new A.RgbColorModelHex { Val = "A5A5A5" }),
                new A.Accent4Color(new A.RgbColorModelHex { Val = "FFC000" }),
                new A.Accent5Color(new A.RgbColorModelHex { Val = "5B9BD5" }),
                new A.Accent6Color(new A.RgbColorModelHex { Val = "70AD47" }),
                new A.Hyperlink(new A.RgbColorModelHex { Val = "0563C1" }),
                new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = "954F72" })) { Name = "Office" },
            new A.FontScheme(
                new A.MajorFont(new A.LatinFont { Typeface = "Arial" }, new A.EastAsianFont { Typeface = string.Empty }, new A.ComplexScriptFont { Typeface = string.Empty }),
                new A.MinorFont(new A.LatinFont { Typeface = "Arial" }, new A.EastAsianFont { Typeface = string.Empty }, new A.ComplexScriptFont { Typeface = string.Empty })) { Name = "Office" },
            new A.FormatScheme(
                new A.FillStyleList(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }), new A.GradientFill(), new A.GradientFill()),
                new A.LineStyleList(new A.Outline(), new A.Outline(), new A.Outline()),
                new A.EffectStyleList(new A.EffectStyle(new A.EffectList()), new A.EffectStyle(new A.EffectList()), new A.EffectStyle(new A.EffectList())),
                new A.BackgroundFillStyleList(new A.SolidFill(), new A.SolidFill(), new A.SolidFill())) { Name = "Office" }),
        new A.ObjectDefaults(),
        new A.ExtraColorSchemeList());

    private static void UpdateText(NotesSlidePart notesPart, PresentationNoteUpdate update)
    {
        var notesSlide = notesPart.NotesSlide
            ?? throw new OfficeWorkerException("invalid_notes", "备注页缺少 notesSlide 根元素");
        var bodyShape = notesSlide.CommonSlideData?.ShapeTree?.Elements<P.Shape>().FirstOrDefault(shape =>
            shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == P.PlaceholderValues.Body)
            ?? throw new OfficeWorkerException("invalid_notes", "备注页模板缺少正文占位符");
        var textBody = bodyShape.TextBody
            ?? throw new OfficeWorkerException("invalid_notes", "备注页模板缺少文本容器");
        var existing = string.Concat(textBody.Descendants<A.Text>().Select(node => node.Text));
        var finalText = update.Append && !string.IsNullOrWhiteSpace(existing)
            ? existing.Trim() + Environment.NewLine + update.Text
            : update.Text;
        textBody.RemoveAllChildren<A.Paragraph>();
        textBody.Append(new A.Paragraph(new A.Run(
            new A.RunProperties { Language = "zh-CN" },
            new A.Text(finalText))));
        notesSlide.Save(notesPart);
    }
}
