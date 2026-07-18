using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Office;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Presentation;

internal sealed class PresentationService(OfficeApplicationProvider applications, OfficeDocumentService documents)
{
    private static readonly string[] ProgIds = ["PowerPoint.Application", "Wpp.Application", "Kwpp.Application"];
    private string? preferredProgId;
    private string? activePresentationPath;

    public object DetectStatus()
    {
        return documents.DetectStatus("presentation");
    }

    public object Open(string filePath)
    {
        var fullPath = Path.GetFullPath(filePath);
        if (!File.Exists(fullPath)) throw new OfficeWorkerException("file_not_found", $"演示文稿不存在: {fullPath}");
        using var handle = applications.GetOrCreate(OrderedProgIds(), "未找到可用的 PowerPoint/WPS 演示 COM 应用");
        dynamic app = handle.Application;
        object? presentations = null;
        object? presentation = null;
        try
        {
            app.Visible = -1;
            presentations = app.Presentations;
            dynamic presentationsApi = presentations;
            presentation = presentationsApi.Open(fullPath);
            dynamic presentationApi = presentation;
            preferredProgId = handle.ProgId;
            activePresentationPath = Convert.ToString(presentationApi.FullName) ?? fullPath;
            return new { success = true, presentationName = Convert.ToString(presentationApi.Name) };
        }
        finally
        {
            ComInterop.Release(presentation);
            ComInterop.Release(presentations);
        }
    }

    public object Inspect()
    {
        return WithPresentation((app, presentation) =>
        {
            dynamic pres = presentation;
            object? slides = null;
            try
            {
                slides = pres.Slides;
                dynamic slidesApi = slides;
                var items = new List<object>();
                for (var index = 1; index <= Convert.ToInt32(slidesApi.Count); index++)
                {
                    object? slide = null;
                    object? shapes = null;
                    try
                    {
                        slide = slidesApi.Item(index);
                        dynamic slideApi = slide;
                        shapes = slideApi.Shapes;
                        dynamic shapesApi = shapes;
                        items.Add(new
                        {
                            index = Convert.ToInt32(slideApi.SlideIndex),
                            name = Convert.ToString(slideApi.Name),
                            shapeCount = Convert.ToInt32(shapesApi.Count),
                            textShapes = ReadShapeTexts(shapesApi),
                        });
                    }
                    finally
                    {
                        ComInterop.Release(shapes);
                        ComInterop.Release(slide);
                    }
                }

                return new
                {
                    app = Convert.ToString(app.Name),
                    progId = preferredProgId,
                    name = Convert.ToString(pres.Name),
                    path = Convert.ToString(pres.FullName),
                    slideCount = Convert.ToInt32(slidesApi.Count),
                    slides = items,
                };
            }
            finally
            {
                ComInterop.Release(slides);
            }
        });
    }

    public object ReadSlide(int slideIndex)
    {
        return WithPresentation((_, presentation) =>
        {
            dynamic pres = presentation;
            object? slides = null;
            object? slide = null;
            object? shapes = null;
            try
            {
                slides = pres.Slides;
                dynamic slidesApi = slides;
                slide = slidesApi.Item(Math.Max(1, slideIndex));
                dynamic slideApi = slide;
                shapes = slideApi.Shapes;
                dynamic shapesApi = shapes;
                return new
                {
                    index = Convert.ToInt32(slideApi.SlideIndex),
                    name = Convert.ToString(slideApi.Name),
                    shapeCount = Convert.ToInt32(shapesApi.Count),
                    textShapes = ReadShapeTexts(shapesApi),
                };
            }
            finally
            {
                ComInterop.Release(shapes);
                ComInterop.Release(slide);
                ComInterop.Release(slides);
            }
        });
    }

    public object AddSlide(string? title, string? body, string layout)
    {
        return WithPresentation((_, presentation) =>
        {
            dynamic pres = presentation;
            object? slides = null;
            object? slide = null;
            try
            {
                slides = pres.Slides;
                dynamic slidesApi = slides;
                var index = Convert.ToInt32(slidesApi.Count) + 1;
                slide = slidesApi.Add(index, ResolveLayout(layout));
                dynamic slideApi = slide;
                if (!string.IsNullOrEmpty(title)) SetPlaceholderText(slideApi, 1, title);
                if (!string.IsNullOrEmpty(body)) SetPlaceholderText(slideApi, 2, body);
                return new { index, title = title ?? string.Empty, layout };
            }
            finally
            {
                ComInterop.Release(slide);
                ComInterop.Release(slides);
            }
        });
    }

    public object SetShapeText(int slideIndex, string text, string? shapeName, int shapeIndex)
    {
        return WithPresentation((_, presentation) =>
        {
            dynamic pres = presentation;
            object? slides = null;
            object? slide = null;
            object? shapes = null;
            object? shape = null;
            object? textFrame = null;
            object? textRange = null;
            try
            {
                slides = pres.Slides;
                dynamic slidesApi = slides;
                slide = slidesApi.Item(Math.Max(1, slideIndex));
                dynamic slideApi = slide;
                shapes = slideApi.Shapes;
                dynamic shapesApi = shapes;
                shape = string.IsNullOrWhiteSpace(shapeName)
                    ? shapesApi.Item(Math.Max(1, shapeIndex))
                    : shapesApi.Item(shapeName);
                dynamic shapeApi = shape;
                textFrame = shapeApi.TextFrame;
                dynamic textFrameApi = textFrame;
                textRange = textFrameApi.TextRange;
                dynamic textRangeApi = textRange;
                textRangeApi.Text = text;
                return new { slideIndex, shapeName = Convert.ToString(shapeApi.Name), characters = text.Length };
            }
            finally
            {
                ComInterop.Release(textRange);
                ComInterop.Release(textFrame);
                ComInterop.Release(shape);
                ComInterop.Release(shapes);
                ComInterop.Release(slide);
                ComInterop.Release(slides);
            }
        });
    }

    public object ReplaceText(string findText, string replaceText, bool matchCase)
    {
        return WithPresentation((_, presentation) =>
        {
            dynamic pres = presentation;
            object? slides = null;
            var replacements = 0;
            try
            {
                slides = pres.Slides;
                dynamic slidesApi = slides;
                for (var slideIndex = 1; slideIndex <= Convert.ToInt32(slidesApi.Count); slideIndex++)
                {
                    object? slide = null;
                    object? shapes = null;
                    try
                    {
                        slide = slidesApi.Item(slideIndex);
                        dynamic slideApi = slide;
                        shapes = slideApi.Shapes;
                        dynamic shapesApi = shapes;
                        for (var shapeIndex = 1; shapeIndex <= Convert.ToInt32(shapesApi.Count); shapeIndex++)
                        {
                            object? shape = null;
                            object? textFrame = null;
                            object? textRange = null;
                            try
                            {
                                shape = shapesApi.Item(shapeIndex);
                                dynamic shapeApi = shape;
                                if (Convert.ToInt32(shapeApi.HasTextFrame) == 0) continue;
                                textFrame = shapeApi.TextFrame;
                                dynamic textFrameApi = textFrame;
                                if (Convert.ToInt32(textFrameApi.HasText) == 0) continue;
                                textRange = textFrameApi.TextRange;
                                dynamic textRangeApi = textRange;
                                string original = Convert.ToString((object?)textRangeApi.Text) ?? string.Empty;
                                var updated = ReplaceAll(original, findText, replaceText, matchCase, out var count);
                                if (count == 0) continue;
                                textRangeApi.Text = updated;
                                replacements += count;
                            }
                            finally
                            {
                                ComInterop.Release(textRange);
                                ComInterop.Release(textFrame);
                                ComInterop.Release(shape);
                            }
                        }
                    }
                    finally
                    {
                        ComInterop.Release(shapes);
                        ComInterop.Release(slide);
                    }
                }

                return new { replacements };
            }
            finally
            {
                ComInterop.Release(slides);
            }
        });
    }

    public object Save(string? saveAsPath)
    {
        return WithPresentation((_, presentation) =>
        {
            dynamic pres = presentation;
            if (string.IsNullOrWhiteSpace(saveAsPath))
            {
                pres.Save();
            }
            else
            {
                var fullPath = Path.GetFullPath(saveAsPath);
                Directory.CreateDirectory(Path.GetDirectoryName(fullPath) ?? Environment.CurrentDirectory);
                pres.SaveAs(fullPath);
                activePresentationPath = fullPath;
            }

            return new { success = true };
        });
    }

    private object WithPresentation(Func<dynamic, object, object> operation)
    {
        using var handle = applications.GetActiveRequired(OrderedProgIds(), "PowerPoint 或 WPS 演示未运行，请先打开文档");
        preferredProgId = handle.ProgId;
        dynamic app = handle.Application;
        object? presentation = ResolvePresentation(app);
        try
        {
            if (presentation is null) throw new OfficeWorkerException("document_not_found", "当前没有活动演示文稿");
            return operation(app, presentation);
        }
        finally
        {
            ComInterop.Release(presentation);
        }
    }

    private object? ResolvePresentation(dynamic app)
    {
        object? presentations = null;
        try
        {
            presentations = app.Presentations;
            dynamic presentationsApi = presentations;
            if (!string.IsNullOrWhiteSpace(activePresentationPath))
            {
                for (var index = 1; index <= Convert.ToInt32(presentationsApi.Count); index++)
                {
                    object? candidate = presentationsApi.Item(index);
                    try
                    {
                        dynamic candidateApi = candidate!;
                        var candidatePath = Convert.ToString(candidateApi.FullName);
                        if (!string.IsNullOrWhiteSpace(candidatePath) && string.Equals(Path.GetFullPath(candidatePath), Path.GetFullPath(activePresentationPath), StringComparison.OrdinalIgnoreCase))
                        {
                            return candidate;
                        }
                    }
                    catch { }
                    ComInterop.Release(candidate);
                }
            }
            return app.ActivePresentation;
        }
        finally
        {
            ComInterop.Release(presentations);
        }
    }

    private static List<object> ReadShapeTexts(dynamic shapes)
    {
        var result = new List<object>();
        for (var index = 1; index <= Convert.ToInt32(shapes.Count); index++)
        {
            object? shape = null;
            object? textFrame = null;
            object? textRange = null;
            try
            {
                shape = shapes.Item(index);
                dynamic shapeApi = shape;
                if (Convert.ToInt32(shapeApi.HasTextFrame) == 0) continue;
                textFrame = shapeApi.TextFrame;
                dynamic textFrameApi = textFrame;
                if (Convert.ToInt32(textFrameApi.HasText) == 0) continue;
                textRange = textFrameApi.TextRange;
                dynamic textRangeApi = textRange;
                result.Add(new { index, name = Convert.ToString(shapeApi.Name), text = Convert.ToString(textRangeApi.Text) });
            }
            finally
            {
                ComInterop.Release(textRange);
                ComInterop.Release(textFrame);
                ComInterop.Release(shape);
            }
        }
        return result;
    }

    private static void SetPlaceholderText(dynamic slide, int index, string text)
    {
        object? shapes = null;
        object? placeholder = null;
        object? textFrame = null;
        object? textRange = null;
        try
        {
            shapes = slide.Shapes;
            dynamic shapesApi = shapes;
            placeholder = index == 1 ? shapesApi.Title : shapesApi.Placeholders.Item(index);
            dynamic placeholderApi = placeholder;
            textFrame = placeholderApi.TextFrame;
            dynamic textFrameApi = textFrame;
            textRange = textFrameApi.TextRange;
            dynamic textRangeApi = textRange;
            textRangeApi.Text = text;
        }
        catch { }
        finally
        {
            ComInterop.Release(textRange);
            ComInterop.Release(textFrame);
            ComInterop.Release(placeholder);
            ComInterop.Release(shapes);
        }
    }

    private static int ResolveLayout(string layout) => layout.ToLowerInvariant() switch
    {
        "title" => 1,
        "titleonly" => 11,
        "blank" => 12,
        _ => 2,
    };

    private static string ReplaceAll(string source, string find, string replacement, bool matchCase, out int count)
    {
        count = 0;
        if (string.IsNullOrEmpty(find)) return source;
        var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var index = source.IndexOf(find, comparison);
        while (index >= 0)
        {
            source = string.Concat(source.AsSpan(0, index), replacement, source.AsSpan(index + find.Length));
            count++;
            index = source.IndexOf(find, index + replacement.Length, comparison);
        }
        return source;
    }

    private IEnumerable<string> OrderedProgIds() =>
        preferredProgId is null ? ProgIds : [preferredProgId, .. ProgIds.Where(id => id != preferredProgId)];
}
