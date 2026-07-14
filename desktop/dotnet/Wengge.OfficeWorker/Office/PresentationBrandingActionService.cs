using System.Text.Json;
using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Office;

internal sealed class PresentationBrandingActionService(OfficeApplicationProvider applications)
{
    private static readonly HashSet<string> Operations = ["applyMasterBranding", "layoutElements"];

    public static bool Supports(string operation) => Operations.Contains(operation);

    public object Execute(OfficeActionRequest request)
    {
        object data;
        string progId;
        var needsWpsThemeFallback = false;
        using (var context = new PresentationActionContext(applications, request))
        {
            data = request.Operation switch
            {
                "applyMasterBranding" => ApplyMasterBranding(context, request),
                "layoutElements" => LayoutElements(context, request),
                _ => throw new OfficeWorkerException("unsupported_operation", $"不支持的演示文稿品牌操作: {request.Operation}"),
            };
            progId = context.ProgId;
            needsWpsThemeFallback = data is BrandingResult && OfficeHostRouting.IsWps(progId);
            if (data is BrandingResult directBranding && directBranding.ThemeColorFailures.Count > 0 && !needsWpsThemeFallback)
                throw new OfficeWorkerException("partial_failure", "一个或多个 PowerPoint 主题颜色未能应用，未保存本次修改", new { failures = directBranding.ThemeColorFailures });
            context.Save(request);
        }
        if (data is BrandingResult branding && needsWpsThemeFallback)
        {
            var outputPath = string.IsNullOrWhiteSpace(request.OutputPath) ? request.FilePath : request.OutputPath;
            var fallbackApplied = PresentationThemePackageUpdater.Apply(outputPath, request.Param("themeColors"));
            if (branding.ThemeColorFailures.Count > 0 && !fallbackApplied)
                throw new OfficeWorkerException("partial_failure", "WPS 主题颜色 COM 写入失败，包级回退也未能完成", new { failures = branding.ThemeColorFailures });
            data = branding with { ThemePackageFallback = fallbackApplied };
        }
        var detail = request.Operation == "applyMasterBranding"
            ? "已统一母版、版式、字体、Logo、页脚和品牌配色"
            : "已精确编辑、对齐、等距分布并检查元素布局";
        return OfficeActionResults.Done(request, "com", detail, OfficeActionResults.WithProgId(data, progId),
            [new OfficeChange(request.Operation == "applyMasterBranding" ? "presentation-brand" : "presentation-layout", "presentation", detail)]);
    }

    private static object ApplyMasterBranding(PresentationActionContext context, OfficeActionRequest request)
    {
        var templatePath = request.StringParam("templatePath");
        if (templatePath.Length > 0)
        {
            templatePath = Path.GetFullPath(templatePath);
            if (!File.Exists(templatePath)) throw new OfficeWorkerException("file_not_found", $"PPT 模板不存在: {templatePath}");
            context.Presentation.ApplyTemplate(templatePath);
        }

        var updatedMasters = 0;
        var updatedLayouts = 0;
        var themeFailures = new List<object>();
        object? designs = null;
        try
        {
            designs = context.Presentation.Designs;
            dynamic designsApi = designs;
            for (var designIndex = 1; designIndex <= Convert.ToInt32(designsApi.Count); designIndex++)
            {
                object? design = null;
                object? master = null;
                object? masterShapes = null;
                object? layouts = null;
                try
                {
                    design = designsApi.Item(designIndex);
                    master = ((dynamic)design).SlideMaster;
                    ApplyMasterSettings(context, master, request, themeFailures);
                    masterShapes = ((dynamic)master).Shapes;
                    ApplyShapeCollection(masterShapes, request);
                    AddOrReplaceLogo(context, master, request);
                    layouts = ((dynamic)master).CustomLayouts;
                    dynamic layoutsApi = layouts;
                    for (var layoutIndex = 1; layoutIndex <= Convert.ToInt32(layoutsApi.Count); layoutIndex++)
                    {
                        object? layout = null;
                        object? shapes = null;
                        try
                        {
                            layout = layoutsApi.Item(layoutIndex);
                            shapes = ((dynamic)layout).Shapes;
                            ApplyShapeCollection(shapes, request);
                            updatedLayouts++;
                        }
                        finally
                        {
                            ComInterop.Release(shapes);
                            ComInterop.Release(layout);
                        }
                    }
                    updatedMasters++;
                }
                finally
                {
                    ComInterop.Release(layouts);
                    ComInterop.Release(masterShapes);
                    ComInterop.Release(master);
                    ComInterop.Release(design);
                }
            }
        }
        finally
        {
            ComInterop.Release(designs);
        }

        var updatedSlides = ApplyBrandingToSlides(context, request);
        var theme = PresentationInspectionActionService.InspectTheme(context).Theme;
        return new BrandingResult(
            templatePath.Length > 0 ? templatePath : null,
            new { masters = updatedMasters, layouts = updatedLayouts, slides = updatedSlides },
            themeFailures,
            false,
            theme,
            context.ProgId);
    }

    private static void ApplyMasterSettings(PresentationActionContext context, object master, OfficeActionRequest request, List<object> themeFailures)
    {
        if (request.StringParam("backgroundColor").Length > 0)
        {
            object? background = null;
            object? fill = null;
            object? color = null;
            try
            {
                background = ((dynamic)master).Background;
                fill = ((dynamic)background).Fill;
                ((dynamic)fill).Solid();
                color = ((dynamic)fill).ForeColor;
                ((dynamic)color).RGB = Color(request.Param("backgroundColor"), 0xFFFFFF);
            }
            finally
            {
                ComInterop.Release(color);
                ComInterop.Release(fill);
                ComInterop.Release(background);
            }
        }

        SetHeadersFooters(master, request, setText: true);
        var themeColors = request.Param("themeColors");
        if (themeColors.ValueKind != JsonValueKind.Array) return;
        object? theme = null;
        object? scheme = null;
        try
        {
            theme = ((dynamic)master).Theme;
            scheme = ((dynamic)theme).ThemeColorScheme;
            foreach (var rule in themeColors.EnumerateArray())
            {
                var index = Int(rule, "index");
                if (index is < 1 or > 12 || !rule.TryGetProperty("value", out var value)) continue;
                object? themeColor = null;
                try
                {
                    themeColor = ((dynamic)scheme).Colors(index);
                    ((dynamic)themeColor).RGB = Color(value, ExcelActionService.OleColor(request.StringParam("accentColor", "1F4E79")));
                }
                catch (Exception exception)
                {
                    themeFailures.Add(new { index, error = exception.Message });
                }
                finally
                {
                    ComInterop.Release(themeColor);
                }
            }
        }
        catch (Exception exception)
        {
            themeFailures.Add(new { index = 0, error = exception.Message });
        }
        finally
        {
            ComInterop.Release(scheme);
            ComInterop.Release(theme);
        }
    }

    private static int ApplyBrandingToSlides(PresentationActionContext context, OfficeActionRequest request)
    {
        object? slides = null;
        try
        {
            slides = context.Presentation.Slides;
            dynamic slidesApi = slides;
            var count = Convert.ToInt32(slidesApi.Count);
            for (var slideIndex = 1; slideIndex <= count; slideIndex++)
            {
                object? slide = null;
                object? shapes = null;
                try
                {
                    slide = slidesApi.Item(slideIndex);
                    shapes = ((dynamic)slide).Shapes;
                    ApplyShapeCollection(shapes, request);
                    SetHeadersFooters(slide, request, setText: false);
                    ApplyMappedLayout(slide, request);
                }
                finally
                {
                    ComInterop.Release(shapes);
                    ComInterop.Release(slide);
                }
            }
            return count;
        }
        finally
        {
            ComInterop.Release(slides);
        }
    }

    private static void ApplyShapeCollection(object shapes, OfficeActionRequest request)
    {
        dynamic shapesApi = shapes;
        for (var index = 1; index <= Convert.ToInt32(shapesApi.Count); index++)
        {
            object? shape = null;
            try
            {
                shape = shapesApi.Item(index);
                ApplyShapeBrand(shape, request);
            }
            finally
            {
                ComInterop.Release(shape);
            }
        }
    }

    private static void ApplyShapeBrand(object shape, OfficeActionRequest request)
    {
        if (Safe(() => Convert.ToInt32(((dynamic)shape).Type), 0) == 6)
        {
            object? groupItems = null;
            try
            {
                groupItems = ((dynamic)shape).GroupItems;
                dynamic groupApi = groupItems;
                for (var index = 1; index <= Convert.ToInt32(groupApi.Count); index++)
                {
                    object? item = null;
                    try { item = groupApi.Item(index); ApplyShapeBrand(item, request); }
                    finally { ComInterop.Release(item); }
                }
            }
            catch
            {
                // Group traversal differs between Office and WPS; top-level formatting still applies.
            }
            finally
            {
                ComInterop.Release(groupItems);
            }
        }

        object? frame = null;
        object? range = null;
        object? font = null;
        object? fontColor = null;
        try
        {
            if (Convert.ToInt32(((dynamic)shape).HasTextFrame) != 0)
            {
                frame = ((dynamic)shape).TextFrame;
                if (Convert.ToInt32(((dynamic)frame).HasText) != 0)
                {
                    range = ((dynamic)frame).TextRange;
                    font = ((dynamic)range).Font;
                    var currentFont = Safe(() => Convert.ToString(((dynamic)font).Name) ?? string.Empty, string.Empty);
                    var targetFont = FontFor(currentFont, request);
                    if (targetFont.Length > 0) ((dynamic)font).Name = targetFont;
                    if (request.BoolParam("applyAccentToText"))
                    {
                        fontColor = ((dynamic)font).Color;
                        ((dynamic)fontColor).RGB = ExcelActionService.OleColor(request.StringParam("accentColor", "1F4E79"));
                    }
                }
            }
        }
        catch
        {
            // Some embedded objects advertise a text frame but reject formatting.
        }
        finally
        {
            ComInterop.Release(fontColor);
            ComInterop.Release(font);
            ComInterop.Release(range);
            ComInterop.Release(frame);
        }
    }

    private static string FontFor(string currentFont, OfficeActionRequest request)
    {
        var map = request.Param("fontMap");
        if (map.ValueKind == JsonValueKind.Object)
        {
            foreach (var mapping in map.EnumerateObject())
                if (string.Equals(mapping.Name, currentFont, StringComparison.OrdinalIgnoreCase) && mapping.Value.ValueKind == JsonValueKind.String)
                    return mapping.Value.GetString() ?? currentFont;
        }
        return request.StringParam("fontName");
    }

    private static void AddOrReplaceLogo(PresentationActionContext context, object master, OfficeActionRequest request)
    {
        var logoPath = request.StringParam("logoPath");
        if (logoPath.Length == 0) return;
        logoPath = Path.GetFullPath(logoPath);
        if (!File.Exists(logoPath)) throw new OfficeWorkerException("file_not_found", $"品牌 Logo 不存在: {logoPath}");
        object? shapes = null;
        object? pageSetup = null;
        object? logo = null;
        try
        {
            shapes = ((dynamic)master).Shapes;
            dynamic shapesApi = shapes;
            for (var index = Convert.ToInt32(shapesApi.Count); index >= 1; index--)
            {
                object? candidate = null;
                object? tags = null;
                try
                {
                    candidate = shapesApi.Item(index);
                    tags = ((dynamic)candidate).Tags;
                    if (string.Equals(Convert.ToString(((dynamic)tags).Item("WENGGE_BRAND_LOGO")), "1", StringComparison.Ordinal))
                        ((dynamic)candidate).Delete();
                }
                catch
                {
                    // Untagged shapes are intentionally preserved.
                }
                finally
                {
                    ComInterop.Release(tags);
                    ComInterop.Release(candidate);
                }
            }
            pageSetup = context.Presentation.PageSetup;
            var width = request.DoubleParam("logoWidth", 100);
            var height = request.Params.TryGetProperty("logoHeight", out _) ? request.DoubleParam("logoHeight") : -1;
            var left = request.Params.TryGetProperty("logoLeft", out _)
                ? request.DoubleParam("logoLeft")
                : Convert.ToDouble(((dynamic)pageSetup).SlideWidth) - width - 24;
            var top = request.Params.TryGetProperty("logoTop", out _) ? request.DoubleParam("logoTop") : 18;
            logo = shapesApi.AddPicture(logoPath, false, true, left, top, width, height);
            ((dynamic)logo).Name = "Wengge Brand Logo";
            object? logoTags = null;
            try { logoTags = ((dynamic)logo).Tags; ((dynamic)logoTags).Add("WENGGE_BRAND_LOGO", "1"); }
            finally { ComInterop.Release(logoTags); }
            ((dynamic)logo).LockAspectRatio = -1;
        }
        finally
        {
            ComInterop.Release(logo);
            ComInterop.Release(pageSetup);
            ComInterop.Release(shapes);
        }
    }

    private static void SetHeadersFooters(object owner, OfficeActionRequest request, bool setText)
    {
        object? headersFooters = null;
        object? footer = null;
        object? slideNumber = null;
        try
        {
            headersFooters = ((dynamic)owner).HeadersFooters;
            footer = ((dynamic)headersFooters).Footer;
            slideNumber = ((dynamic)headersFooters).SlideNumber;
            var footerText = request.StringParam("footerText");
            if (footerText.Length > 0)
            {
                ((dynamic)footer).Visible = -1;
                if (setText) ((dynamic)footer).Text = footerText;
            }
            ((dynamic)slideNumber).Visible = request.BoolParam("showSlideNumber", true) ? -1 : 0;
        }
        catch
        {
            // Headers and footers are not exposed consistently by WPS layouts.
        }
        finally
        {
            ComInterop.Release(slideNumber);
            ComInterop.Release(footer);
            ComInterop.Release(headersFooters);
        }
    }

    private static void ApplyMappedLayout(object slide, OfficeActionRequest request)
    {
        var mappings = request.Param("layoutMap");
        if (mappings.ValueKind != JsonValueKind.Array) return;
        dynamic slideApi = slide;
        var slideIndex = Safe(() => Convert.ToInt32(slideApi.SlideIndex), 0);
        var slideName = Safe(() => Convert.ToString(slideApi.Name) ?? string.Empty, string.Empty);
        string? layoutName = null;
        foreach (var mapping in mappings.EnumerateArray())
        {
            if ((Int(mapping, "slideIndex") == slideIndex && slideIndex > 0)
                || string.Equals(String(mapping, "slideName"), slideName, StringComparison.OrdinalIgnoreCase))
            {
                layoutName = String(mapping, "layoutName");
                break;
            }
        }
        if (string.IsNullOrWhiteSpace(layoutName)) return;
        object? design = null;
        object? master = null;
        object? layouts = null;
        try
        {
            design = slideApi.Design;
            master = ((dynamic)design).SlideMaster;
            layouts = ((dynamic)master).CustomLayouts;
            dynamic layoutsApi = layouts;
            for (var index = 1; index <= Convert.ToInt32(layoutsApi.Count); index++)
            {
                object? layout = null;
                try
                {
                    layout = layoutsApi.Item(index);
                    if (string.Equals(Convert.ToString(((dynamic)layout).Name), layoutName, StringComparison.OrdinalIgnoreCase))
                    {
                        slideApi.CustomLayout = layout;
                        return;
                    }
                }
                finally
                {
                    ComInterop.Release(layout);
                }
            }
        }
        finally
        {
            ComInterop.Release(layouts);
            ComInterop.Release(master);
            ComInterop.Release(design);
        }
    }

    private static object LayoutElements(PresentationActionContext context, OfficeActionRequest request)
    {
        var slideIndexes = request.BoolParam("allSlides")
            ? Enumerable.Range(1, SlideCount(context)).ToArray()
            : [request.SlideIndex()];
        var selectorFailures = PreflightEditSelectors(context, slideIndexes, request);
        if (selectorFailures.Count > 0)
            throw new OfficeWorkerException("partial_failure", "一个或多个 PPT 编辑目标不存在，未执行任何布局修改", new { failures = selectorFailures });
        var failures = new List<object>();
        var edited = 0;
        foreach (var slideIndex in slideIndexes)
        {
            object? slide = null;
            List<object>? items = null;
            try
            {
                slide = context.GetSlide(slideIndex);
                edited += ApplyPreciseEdits(slide, request, failures);
                items = SelectShapes(slide, request);
                if (request.StringParam("mode", "grid") is "grid" or "auto") ApplyGrid(context, items, request);
                if (request.StringParam("align").Length > 0) Align(items, request.StringParam("align"));
                if (request.StringParam("distribute").Length > 0) Distribute(items, request.StringParam("distribute"));
                if (request.BoolParam("fitToSlide")) FitToSlide(context, items);
            }
            finally
            {
                if (items is not null) foreach (var item in items) ComInterop.Release(item);
                ComInterop.Release(slide);
            }
        }
        if (failures.Count > 0)
            throw new OfficeWorkerException("partial_failure", "一个或多个 PPT 编辑未能应用，演示文稿未保存", new { failures });
        var inspection = PresentationInspectionActionService.InspectSlides(context, request);
        return new { editedShapes = edited, editFailures = failures, inspection };
    }

    private static List<object> PreflightEditSelectors(PresentationActionContext context, IReadOnlyList<int> slideIndexes, OfficeActionRequest request)
    {
        var edits = request.Param("edits");
        if (edits.ValueKind != JsonValueKind.Array) return [];
        var requested = edits.EnumerateArray().Select((edit, ordinal) => new
        {
            Edit = edit,
            Ordinal = ordinal,
            Name = String(edit, "shapeName"),
            Index = Int(edit, "shapeIndex"),
        }).ToArray();
        var matched = new bool[requested.Length];
        foreach (var slideIndex in slideIndexes)
        {
            object? slide = null;
            object? shapes = null;
            try
            {
                slide = context.GetSlide(slideIndex);
                shapes = ((dynamic)slide).Shapes;
                dynamic shapesApi = shapes;
                foreach (var selector in requested)
                {
                    if (matched[selector.Ordinal] || (selector.Name.Length == 0 && selector.Index <= 0)) continue;
                    object? shape = null;
                    try
                    {
                        shape = selector.Name.Length > 0 ? shapesApi.Item(selector.Name) : shapesApi.Item(selector.Index);
                        matched[selector.Ordinal] = shape is not null;
                    }
                    catch { }
                    finally { ComInterop.Release(shape); }
                }
            }
            finally
            {
                ComInterop.Release(shapes);
                ComInterop.Release(slide);
            }
        }
        return requested.Where(selector => !matched[selector.Ordinal]).Select(selector => (object)new
        {
            editIndex = selector.Ordinal,
            shapeName = selector.Name,
            shapeIndex = selector.Index,
            error = selector.Name.Length == 0 && selector.Index <= 0 ? "编辑项缺少 shapeName 或 shapeIndex" : "在所选幻灯片中找不到目标形状",
        }).ToList();
    }

    private static int ApplyPreciseEdits(object slide, OfficeActionRequest request, List<object> failures)
    {
        var edits = request.Param("edits");
        if (edits.ValueKind != JsonValueKind.Array) return 0;
        object? shapes = null;
        try
        {
            shapes = ((dynamic)slide).Shapes;
            dynamic shapesApi = shapes;
            var edited = 0;
            foreach (var edit in edits.EnumerateArray())
            {
                object? shape = null;
                try
                {
                    var name = String(edit, "shapeName");
                    var index = Int(edit, "shapeIndex");
                    try { shape = name.Length > 0 ? shapesApi.Item(name) : index > 0 ? shapesApi.Item(index) : null; }
                    catch { shape = null; }
                    if (shape is null) continue;
                    var changed = false;
                    changed |= TrySet(shape, edit, "preserveAspectRatio", "lockAspectRatio", value => ((dynamic)shape).LockAspectRatio = (bool)value ? -1 : 0, failures);
                    changed |= TrySet(shape, edit, "left", "left", value => ((dynamic)shape).Left = Convert.ToSingle(value), failures);
                    changed |= TrySet(shape, edit, "top", "top", value => ((dynamic)shape).Top = Convert.ToSingle(value), failures);
                    changed |= TrySet(shape, edit, "width", "width", value => ((dynamic)shape).Width = Convert.ToSingle(value), failures);
                    changed |= TrySet(shape, edit, "height", "height", value => ((dynamic)shape).Height = Convert.ToSingle(value), failures);
                    changed |= TrySet(shape, edit, "rotation", "rotation", value => ((dynamic)shape).Rotation = Convert.ToSingle(value), failures);
                    changed |= ApplyTextEdit(shape, edit, failures);
                    changed |= ApplyTableEdits(shape, edit, failures);
                    changed |= ApplyChartEdit(shape, edit, failures);
                    changed |= ApplyCrop(shape, edit, failures);
                    if (changed) edited++;
                }
                finally
                {
                    ComInterop.Release(shape);
                }
            }
            return edited;
        }
        finally
        {
            ComInterop.Release(shapes);
        }
    }

    private static bool TrySet(object shape, JsonElement edit, string jsonName, string propertyName, Action<object> setter, List<object> failures)
    {
        if (!edit.TryGetProperty(jsonName, out var value)) return false;
        object raw = value.ValueKind switch
        {
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Number => value.GetDouble(),
            _ => value.ToString(),
        };
        try { setter(raw); return true; }
        catch (Exception exception)
        {
            failures.Add(new { shape = Safe(() => Convert.ToString(((dynamic)shape).Name) ?? string.Empty, string.Empty), property = propertyName, error = exception.Message });
            return false;
        }
    }

    private static bool ApplyTextEdit(object shape, JsonElement edit, List<object> failures)
    {
        if (!edit.TryGetProperty("text", out _) && !edit.TryGetProperty("fontName", out _) && !edit.TryGetProperty("fontSize", out _)) return false;
        object? frame = null;
        object? range = null;
        object? font = null;
        try
        {
            frame = ((dynamic)shape).TextFrame;
            range = ((dynamic)frame).TextRange;
            if (edit.TryGetProperty("text", out var text)) ((dynamic)range).Text = text.ToString();
            if (edit.TryGetProperty("fontName", out var name) || edit.TryGetProperty("fontSize", out _))
            {
                font = ((dynamic)range).Font;
                if (name.ValueKind == JsonValueKind.String) ((dynamic)font).Name = name.GetString();
                if (edit.TryGetProperty("fontSize", out var size) && size.TryGetDouble(out var fontSize)) ((dynamic)font).Size = fontSize;
            }
            return true;
        }
        catch (Exception exception)
        {
            failures.Add(EditFailure(shape, "text", exception));
            return false;
        }
        finally { ComInterop.Release(font); ComInterop.Release(range); ComInterop.Release(frame); }
    }

    private static bool ApplyTableEdits(object shape, JsonElement edit, List<object> failures)
    {
        if (!edit.TryGetProperty("tableCells", out var cells) || cells.ValueKind != JsonValueKind.Array) return false;
        object? table = null;
        var changed = false;
        try
        {
            if (Convert.ToInt32(((dynamic)shape).HasTable) == 0)
            {
                failures.Add(EditFailure(shape, "tableCells", new InvalidOperationException("目标形状不是表格")));
                return false;
            }
            table = ((dynamic)shape).Table;
            foreach (var cellEdit in cells.EnumerateArray())
            {
                object? cell = null; object? cellShape = null; object? frame = null; object? range = null; object? font = null; object? fill = null; object? color = null;
                try
                {
                    cell = ((dynamic)table).Cell(Int(cellEdit, "row"), Int(cellEdit, "column"));
                    cellShape = ((dynamic)cell).Shape; frame = ((dynamic)cellShape).TextFrame; range = ((dynamic)frame).TextRange;
                    if (cellEdit.TryGetProperty("text", out var text)) ((dynamic)range).Text = text.ToString();
                    if (cellEdit.TryGetProperty("fontName", out var name) || cellEdit.TryGetProperty("fontSize", out _))
                    {
                        font = ((dynamic)range).Font;
                        if (name.ValueKind == JsonValueKind.String) ((dynamic)font).Name = name.GetString();
                        if (cellEdit.TryGetProperty("fontSize", out var size) && size.TryGetDouble(out var fontSize)) ((dynamic)font).Size = fontSize;
                    }
                    if (cellEdit.TryGetProperty("fillColor", out var fillValue))
                    {
                        fill = ((dynamic)cellShape).Fill; ((dynamic)fill).Solid(); color = ((dynamic)fill).ForeColor; ((dynamic)color).RGB = Color(fillValue, 0xFFFFFF);
                    }
                    changed = true;
                }
                catch (Exception exception)
                {
                    failures.Add(EditFailure(shape, $"tableCells[{Int(cellEdit, "row")},{Int(cellEdit, "column")}]", exception));
                }
                finally { ComInterop.Release(color); ComInterop.Release(fill); ComInterop.Release(font); ComInterop.Release(range); ComInterop.Release(frame); ComInterop.Release(cellShape); ComInterop.Release(cell); }
            }
        }
        finally { ComInterop.Release(table); }
        return changed;
    }

    private static bool ApplyChartEdit(object shape, JsonElement edit, List<object> failures)
    {
        if (!edit.TryGetProperty("chart", out var chartEdit) || chartEdit.ValueKind != JsonValueKind.Object) return false;
        object? chart = null; object? title = null;
        try
        {
            if (Convert.ToInt32(((dynamic)shape).HasChart) == 0)
            {
                failures.Add(EditFailure(shape, "chart", new InvalidOperationException("目标形状不是图表")));
                return false;
            }
            chart = ((dynamic)shape).Chart;
            if (chartEdit.TryGetProperty("chartType", out var type) && type.TryGetInt32(out var chartType)) ((dynamic)chart).ChartType = chartType;
            if (chartEdit.TryGetProperty("title", out var titleText))
            {
                ((dynamic)chart).HasTitle = true; title = ((dynamic)chart).ChartTitle; ((dynamic)title).Text = titleText.ToString();
            }
            if (chartEdit.TryGetProperty("hasLegend", out var legend) && legend.ValueKind is JsonValueKind.True or JsonValueKind.False) ((dynamic)chart).HasLegend = legend.GetBoolean();
            return true;
        }
        catch (Exception exception)
        {
            failures.Add(EditFailure(shape, "chart", exception));
            return false;
        }
        finally { ComInterop.Release(title); ComInterop.Release(chart); }
    }

    private static bool ApplyCrop(object shape, JsonElement edit, List<object> failures)
    {
        if (!edit.TryGetProperty("crop", out var crop) || crop.ValueKind != JsonValueKind.Object) return false;
        object? format = null;
        try
        {
            format = ((dynamic)shape).PictureFormat;
            if (crop.TryGetProperty("left", out var left)) ((dynamic)format).CropLeft = left.GetDouble();
            if (crop.TryGetProperty("right", out var right)) ((dynamic)format).CropRight = right.GetDouble();
            if (crop.TryGetProperty("top", out var top)) ((dynamic)format).CropTop = top.GetDouble();
            if (crop.TryGetProperty("bottom", out var bottom)) ((dynamic)format).CropBottom = bottom.GetDouble();
            return true;
        }
        catch (Exception exception)
        {
            failures.Add(EditFailure(shape, "crop", exception));
            return false;
        }
        finally { ComInterop.Release(format); }
    }

    private static object EditFailure(object shape, string property, Exception exception) => new
    {
        shape = Safe(() => Convert.ToString(((dynamic)shape).Name) ?? string.Empty, string.Empty),
        property,
        error = exception.Message,
    };

    private static List<object> SelectShapes(object slide, OfficeActionRequest request)
    {
        var selectedNames = StringArray(request.Param("shapeNames")).ToHashSet(StringComparer.OrdinalIgnoreCase);
        object? shapes = null;
        var result = new List<object>();
        try
        {
            shapes = ((dynamic)slide).Shapes;
            dynamic shapesApi = shapes;
            for (var index = 1; index <= Convert.ToInt32(shapesApi.Count); index++)
            {
                object? shape = null;
                try
                {
                    shape = shapesApi.Item(index);
                    var name = Safe(() => Convert.ToString(((dynamic)shape).Name) ?? string.Empty, string.Empty);
                    if (selectedNames.Count > 0 && !selectedNames.Contains(name)) continue;
                    if (request.BoolParam("excludePlaceholders", true) && IsTitlePlaceholder(shape)) continue;
                    result.Add(shape); shape = null;
                }
                finally { ComInterop.Release(shape); }
            }
            return result;
        }
        finally { ComInterop.Release(shapes); }
    }

    private static bool IsTitlePlaceholder(object shape)
    {
        object? placeholder = null;
        try
        {
            if (Convert.ToInt32(((dynamic)shape).Type) != 14) return false;
            placeholder = ((dynamic)shape).PlaceholderFormat;
            return Convert.ToInt32(((dynamic)placeholder).Type) == 1;
        }
        catch { return false; }
        finally { ComInterop.Release(placeholder); }
    }

    private static void ApplyGrid(PresentationActionContext context, List<object> items, OfficeActionRequest request)
    {
        object? pageSetup = null;
        try
        {
            pageSetup = context.Presentation.PageSetup;
            var columns = Math.Max(1, request.IntParam("columns", 2));
            var margin = request.DoubleParam("margin", 40);
            var gap = request.DoubleParam("gap", 16);
            var cellWidth = (Convert.ToDouble(((dynamic)pageSetup).SlideWidth) - 2 * margin - (columns - 1) * gap) / columns;
            var rowHeight = request.DoubleParam("rowHeight", 140);
            for (var index = 0; index < items.Count; index++)
            {
                dynamic item = items[index];
                var column = index % columns; var row = index / columns;
                item.Left = margin + column * (cellWidth + gap); item.Top = margin + row * (rowHeight + gap);
                if (!request.BoolParam("resize", true)) continue;
                if (request.BoolParam("preserveAspectRatio")) item.LockAspectRatio = -1;
                item.Width = cellWidth;
                if (Convert.ToDouble(item.Height) > rowHeight) item.Height = rowHeight;
            }
        }
        finally { ComInterop.Release(pageSetup); }
    }

    private static void Align(List<object> items, string alignment)
    {
        if (items.Count == 0) return;
        var minLeft = items.Min(item => Convert.ToDouble(((dynamic)item).Left));
        var maxRight = items.Max(item => Convert.ToDouble(((dynamic)item).Left) + Convert.ToDouble(((dynamic)item).Width));
        var minTop = items.Min(item => Convert.ToDouble(((dynamic)item).Top));
        var maxBottom = items.Max(item => Convert.ToDouble(((dynamic)item).Top) + Convert.ToDouble(((dynamic)item).Height));
        foreach (dynamic item in items)
        {
            if (alignment == "left") item.Left = minLeft;
            else if (alignment == "center") item.Left = (minLeft + maxRight - Convert.ToDouble(item.Width)) / 2;
            else if (alignment == "right") item.Left = maxRight - Convert.ToDouble(item.Width);
            else if (alignment == "top") item.Top = minTop;
            else if (alignment == "middle") item.Top = (minTop + maxBottom - Convert.ToDouble(item.Height)) / 2;
            else if (alignment == "bottom") item.Top = maxBottom - Convert.ToDouble(item.Height);
        }
    }

    private static void Distribute(List<object> items, string direction)
    {
        if (items.Count < 3) return;
        var vertical = direction == "vertical";
        var ordered = vertical
            ? items.OrderBy(item => Convert.ToDouble(((dynamic)item).Top)).ToArray()
            : items.OrderBy(item => Convert.ToDouble(((dynamic)item).Left)).ToArray();
        var first = vertical ? Convert.ToDouble(((dynamic)ordered[0]).Top) : Convert.ToDouble(((dynamic)ordered[0]).Left);
        var lastEnd = vertical
            ? Convert.ToDouble(((dynamic)ordered[^1]).Top) + Convert.ToDouble(((dynamic)ordered[^1]).Height)
            : Convert.ToDouble(((dynamic)ordered[^1]).Left) + Convert.ToDouble(((dynamic)ordered[^1]).Width);
        var totalSize = ordered.Sum(item => vertical ? Convert.ToDouble(((dynamic)item).Height) : Convert.ToDouble(((dynamic)item).Width));
        var gap = (lastEnd - first - totalSize) / (ordered.Length - 1);
        var cursor = first;
        foreach (dynamic item in ordered)
        {
            if (vertical) { item.Top = cursor; cursor += Convert.ToDouble(item.Height) + gap; }
            else { item.Left = cursor; cursor += Convert.ToDouble(item.Width) + gap; }
        }
    }

    private static void FitToSlide(PresentationActionContext context, List<object> items)
    {
        object? pageSetup = null;
        try
        {
            pageSetup = context.Presentation.PageSetup;
            var width = Convert.ToDouble(((dynamic)pageSetup).SlideWidth);
            var height = Convert.ToDouble(((dynamic)pageSetup).SlideHeight);
            foreach (dynamic item in items)
            {
                if (Convert.ToDouble(item.Left) < 0) item.Left = 0;
                if (Convert.ToDouble(item.Top) < 0) item.Top = 0;
                if (Convert.ToDouble(item.Left) + Convert.ToDouble(item.Width) > width) item.Left = Math.Max(0, width - Convert.ToDouble(item.Width));
                if (Convert.ToDouble(item.Top) + Convert.ToDouble(item.Height) > height) item.Top = Math.Max(0, height - Convert.ToDouble(item.Height));
            }
        }
        finally { ComInterop.Release(pageSetup); }
    }

    private static int SlideCount(PresentationActionContext context)
    {
        object? slides = null;
        try { slides = context.Presentation.Slides; return Convert.ToInt32(((dynamic)slides).Count); }
        finally { ComInterop.Release(slides); }
    }

    private static int Color(JsonElement value, int fallback)
    {
        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var number)) return number;
        if (value.ValueKind != JsonValueKind.String) return fallback;
        var text = (value.GetString() ?? string.Empty).Trim().TrimStart('#');
        return text.Length == 6 ? ExcelActionService.OleColor(text) : fallback;
    }

    private static string String(JsonElement value, string name, string fallback = "") =>
        value.ValueKind == JsonValueKind.Object && value.TryGetProperty(name, out var property) && property.ValueKind == JsonValueKind.String ? property.GetString() ?? fallback : fallback;

    private static int Int(JsonElement value, string name, int fallback = 0) =>
        value.ValueKind == JsonValueKind.Object && value.TryGetProperty(name, out var property) && property.TryGetInt32(out var result) ? result : fallback;

    private static IEnumerable<string> StringArray(JsonElement value)
    {
        if (value.ValueKind != JsonValueKind.Array) yield break;
        foreach (var item in value.EnumerateArray()) if (item.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(item.GetString())) yield return item.GetString()!;
    }

    private sealed record BrandingResult(
        string? AppliedTemplate,
        object Updated,
        List<object> ThemeColorFailures,
        bool ThemePackageFallback,
        object Theme,
        string ProgId);

    private static T Safe<T>(Func<T> value, T fallback)
    {
        try { return value(); }
        catch { return fallback; }
    }
}
