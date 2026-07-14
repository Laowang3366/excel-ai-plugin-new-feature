using System.Text.Json;
using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Office;

internal sealed class PresentationEditActionService(OfficeApplicationProvider applications)
{
    private static readonly HashSet<string> Operations =
    ["addSlide", "addSlides", "appendSlide", "appendSlides", "addSlideContent", "applyTheme", "deleteSlides", "normalizeLayouts", "insertChart", "insertTable", "replacePictureSlot", "alignShapes", "snapshot"];

    public static bool Supports(string operation) => Operations.Contains(operation);

    public object Execute(OfficeActionRequest request)
    {
        using var context = new PresentationActionContext(applications, request);
        if (request.Operation == "snapshot")
        {
            var output = request.OutputPath ?? Path.Combine(Path.GetDirectoryName(request.FilePath) ?? Environment.CurrentDirectory, "snapshot.png");
            output = Path.GetFullPath(output); Directory.CreateDirectory(Path.GetDirectoryName(output) ?? Environment.CurrentDirectory);
            object? slide = context.GetSlide(request.SlideIndex());
            try { ((dynamic)slide).Export(output, "PNG"); }
            finally { ComInterop.Release(slide); }
            return OfficeActionResults.Done(request, "com", "已导出幻灯片快照",
                OfficeActionResults.WithProgId(new { outputPath = output }, context.ProgId),
                [new OfficeChange("snapshot", output, "已导出幻灯片快照")], output);
        }
        var data = request.Operation switch
        {
            "addSlide" or "appendSlide" => AddSlide(context, request, request.Params),
            "addSlides" or "appendSlides" => AddSlides(context, request),
            "addSlideContent" => AddSlideContent(context, request),
            "deleteSlides" => DeleteSlides(context, request),
            "insertChart" => InsertChart(context, request),
            "insertTable" => InsertTable(context, request),
            "replacePictureSlot" => ReplacePicture(context, request),
            "alignShapes" => AlignShapes(context, request),
            "normalizeLayouts" => NormalizeLayouts(context, request),
            "applyTheme" => ApplyTheme(context, request),
            _ => throw new OfficeWorkerException("unsupported_operation", $"不支持的演示文稿操作: {request.Operation}"),
        };
        context.Save(request);
        return OfficeActionResults.Done(request, "com", "已更新演示文稿", OfficeActionResults.WithProgId(data, context.ProgId),
            [new OfficeChange("presentation-edit", request.Target, "已更新演示文稿")]);
    }

    private static object AddSlides(PresentationActionContext context, OfficeActionRequest request)
    {
        var definitions = request.Param("slides");
        if (definitions.ValueKind != JsonValueKind.Array) throw new OfficeWorkerException("invalid_params", "addSlides 需要 params.slides");
        var slides = new List<object>();
        foreach (var definition in definitions.EnumerateArray()) slides.Add(AddSlide(context, request, definition));
        return new { slides, added = slides.Count };
    }

    private static object AddSlide(PresentationActionContext context, OfficeActionRequest request, JsonElement definition)
    {
        object? slides = null; object? slide = null;
        try
        {
            slides = context.Presentation.Slides; dynamic slidesApi = slides;
            var index = definition.TryGetProperty("index", out var indexValue) && indexValue.TryGetInt32(out var requestedIndex)
                ? Math.Clamp(requestedIndex, 1, Convert.ToInt32(slidesApi.Count) + 1)
                : Convert.ToInt32(slidesApi.Count) + 1;
            var layout = String(definition, "layout", request.StringParam("layout", "titleAndContent"));
            slide = slidesApi.Add(index, Layout(layout)); dynamic slideApi = slide;
            SetPlaceholder(slideApi, 1, String(definition, "title", request.StringParam("title")));
            SetPlaceholder(slideApi, 2, String(definition, "body", request.StringParam("body")));
            return new { index = Convert.ToInt32(slideApi.SlideIndex), layout };
        }
        finally { ComInterop.Release(slide); ComInterop.Release(slides); }
    }

    private static object AddSlideContent(PresentationActionContext context, OfficeActionRequest request)
    {
        object? slide = context.GetSlide(request.SlideIndex()); object? shapes = null; object? shape = null;
        try
        {
            dynamic slideApi = slide; shapes = slideApi.Shapes; dynamic shapesApi = shapes;
            var type = request.StringParam("contentType", request.StringParam("type", "text"));
            if (type == "image")
            {
                var imagePath = Path.GetFullPath(request.StringParam("imagePath"));
                if (!File.Exists(imagePath)) throw new OfficeWorkerException("file_not_found", $"图片不存在: {imagePath}");
                shape = shapesApi.AddPicture(imagePath, false, true, request.DoubleParam("left", 80), request.DoubleParam("top", 120), request.DoubleParam("width", 520), request.DoubleParam("height", 300));
            }
            else
            {
                shape = shapesApi.AddTextbox(1, request.DoubleParam("left", 80), request.DoubleParam("top", 120), request.DoubleParam("width", 520), request.DoubleParam("height", 120));
                ((dynamic)shape).TextFrame.TextRange.Text = request.StringParam("text");
            }
            if (request.StringParam("name").Length > 0) ((dynamic)shape).Name = request.StringParam("name");
            return new { slideIndex = request.SlideIndex(), type, name = Safe(() => ((dynamic)shape).Name) };
        }
        finally { ComInterop.Release(shape); ComInterop.Release(shapes); ComInterop.Release(slide); }
    }

    private static object DeleteSlides(PresentationActionContext context, OfficeActionRequest request)
    {
        var indexes = SlideIndexes(request).Distinct().OrderDescending().ToArray();
        if (indexes.Length == 0) throw new OfficeWorkerException("invalid_params", "deleteSlides 需要 params.slides、params.from/to 或 slide target");
        object? slides = null;
        try
        {
            slides = context.Presentation.Slides; dynamic slidesApi = slides;
            foreach (var index in indexes)
            {
                if (Convert.ToInt32(slidesApi.Count) <= 1) throw new OfficeWorkerException("invalid_operation", "至少需要保留一张幻灯片");
                object? slide = null; try { slide = slidesApi.Item(index); ((dynamic)slide).Delete(); } finally { ComInterop.Release(slide); }
            }
            return new { deleted = indexes };
        }
        finally { ComInterop.Release(slides); }
    }

    private static object InsertChart(PresentationActionContext context, OfficeActionRequest request)
    {
        object? slide = context.GetSlide(request.SlideIndex()); object? shapes = null; object? shape = null;
        try
        {
            shapes = ((dynamic)slide).Shapes; dynamic shapesApi = shapes;
            try { shape = shapesApi.AddChart2(201, ChartType(request.StringParam("chartType")), request.DoubleParam("left", 80), request.DoubleParam("top", 120), request.DoubleParam("width", 520), request.DoubleParam("height", 300)); }
            catch { shape = shapesApi.AddChart(ChartType(request.StringParam("chartType")), request.DoubleParam("left", 80), request.DoubleParam("top", 120), request.DoubleParam("width", 520), request.DoubleParam("height", 300)); }
            if (request.StringParam("name").Length > 0) ((dynamic)shape).Name = request.StringParam("name");
            return new { slideIndex = request.SlideIndex(), name = Safe(() => ((dynamic)shape).Name) };
        }
        finally { ComInterop.Release(shape); ComInterop.Release(shapes); ComInterop.Release(slide); }
    }

    private static object InsertTable(PresentationActionContext context, OfficeActionRequest request)
    {
        var values = request.Param("values");
        var rows = Math.Max(1, request.IntParam("rows", values.ValueKind == JsonValueKind.Array ? values.GetArrayLength() : 2));
        var columns = Math.Max(1, request.IntParam("columns", values.ValueKind == JsonValueKind.Array && values.GetArrayLength() > 0 ? values[0].GetArrayLength() : 2));
        object? slide = context.GetSlide(request.SlideIndex()); object? shapes = null; object? shape = null; object? table = null;
        try
        {
            shapes = ((dynamic)slide).Shapes; shape = ((dynamic)shapes).AddTable(rows, columns, request.DoubleParam("left", 80), request.DoubleParam("top", 120), request.DoubleParam("width", 520), request.DoubleParam("height", 220));
            dynamic shapeApi = shape; if (request.StringParam("name").Length > 0) shapeApi.Name = request.StringParam("name"); table = shapeApi.Table; dynamic tableApi = table;
            if (values.ValueKind == JsonValueKind.Array)
            {
                var rowValues = values.EnumerateArray().ToArray();
                for (var row = 0; row < Math.Min(rows, rowValues.Length); row++)
                {
                    var cells = rowValues[row].ValueKind == JsonValueKind.Array ? rowValues[row].EnumerateArray().ToArray() : [];
                    for (var column = 0; column < Math.Min(columns, cells.Length); column++)
                    {
                        object? cell = null; object? cellShape = null; object? frame = null; object? range = null;
                        try { cell = tableApi.Cell(row + 1, column + 1); cellShape = ((dynamic)cell).Shape; frame = ((dynamic)cellShape).TextFrame; range = ((dynamic)frame).TextRange; ((dynamic)range).Text = cells[column].ToString(); }
                        finally { ComInterop.Release(range); ComInterop.Release(frame); ComInterop.Release(cellShape); ComInterop.Release(cell); }
                    }
                }
            }
            return new { slideIndex = request.SlideIndex(), rows, columns, name = Safe(() => shapeApi.Name) };
        }
        finally { ComInterop.Release(table); ComInterop.Release(shape); ComInterop.Release(shapes); ComInterop.Release(slide); }
    }

    private static object ReplacePicture(PresentationActionContext context, OfficeActionRequest request)
    {
        var imagePath = Path.GetFullPath(request.StringParam("imagePath"));
        if (!File.Exists(imagePath)) throw new OfficeWorkerException("file_not_found", $"图片不存在: {imagePath}");
        object? slide = context.GetSlide(request.SlideIndex()); object? shapes = null; object? old = null; object? picture = null;
        try
        {
            shapes = ((dynamic)slide).Shapes; dynamic shapesApi = shapes;
            var shapeName = request.StringParam("shapeName");
            if (shapeName.Length > 0) { try { old = shapesApi.Item(shapeName); ((dynamic)old).Delete(); } catch { } }
            picture = shapesApi.AddPicture(imagePath, false, true, request.DoubleParam("left", 80), request.DoubleParam("top", 120), request.DoubleParam("width", 520), request.DoubleParam("height", 300));
            if (request.StringParam("name").Length > 0) ((dynamic)picture).Name = request.StringParam("name");
            if (request.BoolParam("preserveAspectRatio")) ((dynamic)picture).LockAspectRatio = -1;
            return new { slideIndex = request.SlideIndex(), imagePath, name = Safe(() => ((dynamic)picture).Name) };
        }
        finally { ComInterop.Release(picture); ComInterop.Release(old); ComInterop.Release(shapes); ComInterop.Release(slide); }
    }

    private static object AlignShapes(PresentationActionContext context, OfficeActionRequest request)
    {
        object? slide = context.GetSlide(request.SlideIndex()); object? shapes = null; var aligned = 0;
        try
        {
            shapes = ((dynamic)slide).Shapes; dynamic shapesApi = shapes;
            for (var index = 1; index <= Convert.ToInt32(shapesApi.Count); index++)
            {
                object? shape = null;
                try
                {
                    shape = shapesApi.Item(index); dynamic api = shape;
                    if (Convert.ToDouble(api.Left) < request.DoubleParam("minLeft", 40)) api.Left = request.DoubleParam("minLeft", 40);
                    if (Convert.ToDouble(api.Top) < request.DoubleParam("minTop", 40)) api.Top = request.DoubleParam("minTop", 40);
                    aligned++;
                }
                finally { ComInterop.Release(shape); }
            }
            return new { slideIndex = request.SlideIndex(), aligned };
        }
        finally { ComInterop.Release(shapes); ComInterop.Release(slide); }
    }

    private static object NormalizeLayouts(PresentationActionContext context, OfficeActionRequest request)
    {
        object? slides = null; var normalized = 0;
        try
        {
            slides = context.Presentation.Slides; dynamic slidesApi = slides;
            for (var slideIndex = 1; slideIndex <= Convert.ToInt32(slidesApi.Count); slideIndex++)
            {
                object? slide = null; object? shapes = null;
                try
                {
                    slide = slidesApi.Item(slideIndex); shapes = ((dynamic)slide).Shapes; dynamic shapesApi = shapes;
                    for (var index = 1; index <= Convert.ToInt32(shapesApi.Count); index++)
                    {
                        object? shape = null; try { shape = shapesApi.Item(index); dynamic api = shape; if (Convert.ToDouble(api.Width) > 600) api.Width = 600; if (Convert.ToDouble(api.Height) > 360) api.Height = 360; normalized++; } finally { ComInterop.Release(shape); }
                    }
                }
                finally { ComInterop.Release(shapes); ComInterop.Release(slide); }
            }
            return new { normalized };
        }
        finally { ComInterop.Release(slides); }
    }

    private static object ApplyTheme(PresentationActionContext context, OfficeActionRequest request)
    {
        var color = ExcelActionService.OleColor(request.StringParam("accentColor", "1F4E79"));
        object? slides = null; var updated = 0;
        try
        {
            slides = context.Presentation.Slides; dynamic slidesApi = slides;
            for (var slideIndex = 1; slideIndex <= Convert.ToInt32(slidesApi.Count); slideIndex++)
            {
                object? slide = null; object? shapes = null;
                try
                {
                    slide = slidesApi.Item(slideIndex); shapes = ((dynamic)slide).Shapes; dynamic shapesApi = shapes;
                    for (var index = 1; index <= Convert.ToInt32(shapesApi.Count); index++)
                    {
                        object? shape = null; object? frame = null; object? range = null; object? font = null; object? fontColor = null;
                        try
                        {
                            shape = shapesApi.Item(index); dynamic shapeApi = shape;
                            if (Convert.ToInt32(shapeApi.HasTextFrame) == 0) continue;
                            frame = shapeApi.TextFrame; if (Convert.ToInt32(((dynamic)frame).HasText) == 0) continue;
                            range = ((dynamic)frame).TextRange; font = ((dynamic)range).Font; fontColor = ((dynamic)font).Color; ((dynamic)fontColor).RGB = color; updated++;
                        }
                        finally { ComInterop.Release(fontColor); ComInterop.Release(font); ComInterop.Release(range); ComInterop.Release(frame); ComInterop.Release(shape); }
                    }
                }
                finally { ComInterop.Release(shapes); ComInterop.Release(slide); }
            }
            return new { updated, accentColor = request.StringParam("accentColor", "1F4E79") };
        }
        finally { ComInterop.Release(slides); }
    }

    private static IEnumerable<int> SlideIndexes(OfficeActionRequest request)
    {
        var slides = request.Param("slides");
        if (slides.ValueKind == JsonValueKind.Array) foreach (var value in slides.EnumerateArray()) if (value.TryGetInt32(out var index)) yield return index;
        var from = request.IntParam("from"); var to = request.IntParam("to");
        if (from > 0 && to >= from) for (var index = from; index <= to; index++) yield return index;
        if (request.Target?.StartsWith("slide:", StringComparison.OrdinalIgnoreCase) == true)
        {
            var value = request.Target[6..]; var range = value.Split('-', 2);
            if (range.Length == 2 && int.TryParse(range[0], out from) && int.TryParse(range[1], out to)) for (var index = from; index <= to; index++) yield return index;
            else if (int.TryParse(value, out var index)) yield return index;
        }
    }

    private static void SetPlaceholder(dynamic slide, int index, string text)
    {
        if (text.Length == 0) return;
        object? shapes = null; object? shape = null; object? frame = null; object? range = null;
        try { shapes = slide.Shapes; shape = index == 1 ? ((dynamic)shapes).Title : ((dynamic)shapes).Placeholders.Item(index); frame = ((dynamic)shape).TextFrame; range = ((dynamic)frame).TextRange; ((dynamic)range).Text = text; }
        catch { }
        finally { ComInterop.Release(range); ComInterop.Release(frame); ComInterop.Release(shape); ComInterop.Release(shapes); }
    }

    private static string String(JsonElement value, string name, string fallback = "") => value.TryGetProperty(name, out var property) && property.ValueKind == JsonValueKind.String ? property.GetString() ?? fallback : fallback;
    private static int Layout(string value) => value.ToLowerInvariant() switch { "title" => 1, "titleonly" => 11, "blank" => 12, _ => 2 };
    private static int ChartType(string value) => value.ToLowerInvariant() switch { "line" => 4, "pie" => 5, "bar" => 57, "area" => 1, "scatter" => -4169, _ => 51 };
    private static object? Safe(Func<object?> value) { try { return value(); } catch { return null; } }
}
