using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Office;

internal sealed class PresentationInspectionActionService(OfficeApplicationProvider applications)
{
    private static readonly HashSet<string> Operations =
    [
        "inspectPresentationTheme",
        "inspectSlideElements",
    ];

    public static bool Supports(string operation) => Operations.Contains(operation);

    public object Execute(OfficeActionRequest request)
    {
        using var context = new PresentationActionContext(applications, request);
        var data = request.Operation switch
        {
            "inspectPresentationTheme" => InspectTheme(context),
            "inspectSlideElements" => InspectSlides(context, request),
            _ => throw new OfficeWorkerException("unsupported_operation", $"不支持的演示文稿检查操作: {request.Operation}"),
        };
        return OfficeActionResults.Done(request, "com", "已检查演示文稿", data);
    }

    internal static object InspectSlides(PresentationActionContext context, OfficeActionRequest request)
    {
        var slideIndexes = request.BoolParam("allSlides")
            ? Enumerable.Range(1, SlideCount(context)).ToArray()
            : [request.SlideIndex()];
        var slides = new List<SlideSnapshot>();
        foreach (var index in slideIndexes)
        {
            object? slide = null;
            try
            {
                slide = context.GetSlide(index);
                slides.Add(SnapshotSlide(context, slide));
            }
            finally
            {
                ComInterop.Release(slide);
            }
        }

        return new
        {
            progId = context.ProgId,
            slides,
            summary = new
            {
                slideCount = slides.Count,
                shapeCount = slides.Sum(slide => slide.Shapes.Count),
                overflowCount = slides.Sum(slide => slide.Shapes.Count(shape => shape.TextOverflow)),
                outOfBoundsCount = slides.Sum(slide => slide.Shapes.Count(shape => shape.OutOfBounds)),
                overlapCount = slides.Sum(slide => slide.Overlaps.Count),
            },
        };
    }

    internal static ThemeInspection InspectTheme(PresentationActionContext context)
    {
        object? pageSetup = null;
        object? designs = null;
        object? slideMaster = null;
        object? headersFooters = null;
        object? footer = null;
        object? slideNumber = null;
        try
        {
            pageSetup = context.Presentation.PageSetup;
            designs = context.Presentation.Designs;
            dynamic designsApi = designs;
            var snapshots = new List<object>();
            for (var designIndex = 1; designIndex <= Convert.ToInt32(designsApi.Count); designIndex++)
            {
                object? design = null;
                object? master = null;
                object? layouts = null;
                object? theme = null;
                object? scheme = null;
                try
                {
                    design = designsApi.Item(designIndex);
                    dynamic designApi = design;
                    master = designApi.SlideMaster;
                    dynamic masterApi = master;
                    layouts = masterApi.CustomLayouts;
                    dynamic layoutsApi = layouts;
                    var layoutSnapshots = new List<object>();
                    for (var layoutIndex = 1; layoutIndex <= Convert.ToInt32(layoutsApi.Count); layoutIndex++)
                    {
                        object? layout = null;
                        object? shapes = null;
                        try
                        {
                            layout = layoutsApi.Item(layoutIndex);
                            dynamic layoutApi = layout;
                            shapes = layoutApi.Shapes;
                            layoutSnapshots.Add(new
                            {
                                index = Safe(() => Convert.ToInt32(layoutApi.Index), layoutIndex),
                                name = Safe(() => Convert.ToString(layoutApi.Name) ?? string.Empty, string.Empty),
                                shapeCount = Safe(() => Convert.ToInt32(((dynamic)shapes).Count), 0),
                            });
                        }
                        finally
                        {
                            ComInterop.Release(shapes);
                            ComInterop.Release(layout);
                        }
                    }

                    var colors = new List<object>();
                    try
                    {
                        theme = masterApi.Theme;
                        scheme = ((dynamic)theme).ThemeColorScheme;
                        for (var colorIndex = 1; colorIndex <= 12; colorIndex++)
                        {
                            object? color = null;
                            try
                            {
                                color = ((dynamic)scheme).Colors(colorIndex);
                                colors.Add(new { index = colorIndex, rgb = Convert.ToInt32(((dynamic)color).RGB) });
                            }
                            catch
                            {
                                // Some WPS builds expose only part of the Office theme color collection.
                            }
                            finally
                            {
                                ComInterop.Release(color);
                            }
                        }
                    }
                    catch
                    {
                        // Theme inspection is still useful when the host does not expose theme colors.
                    }

                    object? masterShapes = null;
                    try
                    {
                        masterShapes = masterApi.Shapes;
                        snapshots.Add(new
                        {
                            index = Safe(() => Convert.ToInt32(designApi.Index), designIndex),
                            name = Safe(() => Convert.ToString(designApi.Name) ?? string.Empty, string.Empty),
                            masterName = Safe(() => Convert.ToString(masterApi.Name) ?? string.Empty, string.Empty),
                            masterShapeCount = Safe(() => Convert.ToInt32(((dynamic)masterShapes).Count), 0),
                            layouts = layoutSnapshots,
                            colors,
                        });
                    }
                    finally
                    {
                        ComInterop.Release(masterShapes);
                    }
                }
                finally
                {
                    ComInterop.Release(scheme);
                    ComInterop.Release(theme);
                    ComInterop.Release(layouts);
                    ComInterop.Release(master);
                    ComInterop.Release(design);
                }
            }

            slideMaster = context.Presentation.SlideMaster;
            headersFooters = ((dynamic)slideMaster).HeadersFooters;
            footer = ((dynamic)headersFooters).Footer;
            slideNumber = ((dynamic)headersFooters).SlideNumber;
            return new ThemeInspection(
                new
                {
                    slideCount = SlideCount(context),
                    width = Convert.ToDouble(((dynamic)pageSetup).SlideWidth),
                    height = Convert.ToDouble(((dynamic)pageSetup).SlideHeight),
                    designs = snapshots,
                    footer = new
                    {
                        visible = Safe(() => Convert.ToBoolean(((dynamic)footer).Visible), false),
                        text = Safe(() => Convert.ToString(((dynamic)footer).Text) ?? string.Empty, string.Empty),
                        slideNumberVisible = Safe(() => Convert.ToBoolean(((dynamic)slideNumber).Visible), false),
                    },
                }, context.ProgId);
        }
        finally
        {
            ComInterop.Release(slideNumber);
            ComInterop.Release(footer);
            ComInterop.Release(headersFooters);
            ComInterop.Release(slideMaster);
            ComInterop.Release(designs);
            ComInterop.Release(pageSetup);
        }
    }

    private static SlideSnapshot SnapshotSlide(PresentationActionContext context, object slide)
    {
        object? pageSetup = null;
        object? shapes = null;
        object? customLayout = null;
        try
        {
            pageSetup = context.Presentation.PageSetup;
            var slideWidth = Convert.ToDouble(((dynamic)pageSetup).SlideWidth);
            var slideHeight = Convert.ToDouble(((dynamic)pageSetup).SlideHeight);
            dynamic slideApi = slide;
            shapes = slideApi.Shapes;
            dynamic shapesApi = shapes;
            var snapshots = new List<ShapeSnapshot>();
            for (var index = 1; index <= Convert.ToInt32(shapesApi.Count); index++)
            {
                object? shape = null;
                try
                {
                    shape = shapesApi.Item(index);
                    snapshots.Add(SnapshotShape(shape, slideWidth, slideHeight));
                }
                finally
                {
                    ComInterop.Release(shape);
                }
            }

            var overlaps = new List<object>();
            for (var leftIndex = 0; leftIndex < snapshots.Count; leftIndex++)
            {
                for (var rightIndex = leftIndex + 1; rightIndex < snapshots.Count; rightIndex++)
                {
                    var first = snapshots[leftIndex];
                    var second = snapshots[rightIndex];
                    var width = Math.Min(first.Left + first.Width, second.Left + second.Width) - Math.Max(first.Left, second.Left);
                    var height = Math.Min(first.Top + first.Height, second.Top + second.Height) - Math.Max(first.Top, second.Top);
                    if (width > 1 && height > 1)
                    {
                        overlaps.Add(new { first = first.Name, second = second.Name, width, height, area = width * height });
                    }
                }
            }

            customLayout = Safe<object?>(() => slideApi.CustomLayout, null);
            return new SlideSnapshot(
                Safe(() => Convert.ToInt32(slideApi.SlideIndex), 0),
                Safe(() => Convert.ToString(slideApi.Name) ?? string.Empty, string.Empty),
                customLayout is null ? string.Empty : Safe(() => Convert.ToString(((dynamic)customLayout).Name) ?? string.Empty, string.Empty),
                snapshots,
                overlaps);
        }
        finally
        {
            ComInterop.Release(customLayout);
            ComInterop.Release(shapes);
            ComInterop.Release(pageSetup);
        }
    }

    private static ShapeSnapshot SnapshotShape(object shape, double slideWidth, double slideHeight)
    {
        dynamic api = shape;
        double left = Safe(() => Convert.ToDouble(api.Left), 0d);
        double top = Safe(() => Convert.ToDouble(api.Top), 0d);
        double width = Safe(() => Convert.ToDouble(api.Width), 0d);
        double height = Safe(() => Convert.ToDouble(api.Height), 0d);
        int type = Safe(() => Convert.ToInt32(api.Type), 0);
        var text = ReadShapeText(shape);
        var (boundWidth, boundHeight, overflow) = TextBounds(shape, text, width, height);
        return new ShapeSnapshot(
            Safe(() => Convert.ToInt32(api.Id), 0),
            Safe(() => Convert.ToString(api.Name) ?? string.Empty, string.Empty),
            type,
            ShapeTypeName(type),
            Safe(() => Convert.ToInt32(api.ZOrderPosition), 0),
            left,
            top,
            width,
            height,
            Safe(() => Convert.ToDouble(api.Rotation), 0d),
            text,
            new { width = boundWidth, height = boundHeight },
            overflow,
            left < 0 || top < 0 || left + width > slideWidth + 1 || top + height > slideHeight + 1,
            ReadTable(shape),
            ReadChart(shape),
            type == 13 ? ReadPicture(shape) : null);
    }

    private static string ReadShapeText(object shape)
    {
        object? frame = null;
        object? range = null;
        try
        {
            dynamic api = shape;
            if (Convert.ToInt32(api.HasTextFrame) == 0) return string.Empty;
            frame = api.TextFrame;
            if (Convert.ToInt32(((dynamic)frame).HasText) == 0) return string.Empty;
            range = ((dynamic)frame).TextRange;
            return Convert.ToString(((dynamic)range).Text) ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
        finally
        {
            ComInterop.Release(range);
            ComInterop.Release(frame);
        }
    }

    private static (double Width, double Height, bool Overflow) TextBounds(object shape, string text, double width, double height)
    {
        if (text.Length == 0) return (0, 0, false);
        object? frame = null;
        object? range = null;
        try
        {
            frame = ((dynamic)shape).TextFrame2;
            range = ((dynamic)frame).TextRange;
            var boundWidth = Convert.ToDouble(((dynamic)range).BoundWidth);
            var boundHeight = Convert.ToDouble(((dynamic)range).BoundHeight);
            var availableWidth = Math.Max(1, width - Convert.ToDouble(((dynamic)frame).MarginLeft) - Convert.ToDouble(((dynamic)frame).MarginRight));
            var availableHeight = Math.Max(1, height - Convert.ToDouble(((dynamic)frame).MarginTop) - Convert.ToDouble(((dynamic)frame).MarginBottom));
            return (boundWidth, boundHeight, boundWidth > availableWidth + 1 || boundHeight > availableHeight + 1);
        }
        catch
        {
            return (0, 0, false);
        }
        finally
        {
            ComInterop.Release(range);
            ComInterop.Release(frame);
        }
    }

    private static object? ReadTable(object shape)
    {
        object? table = null;
        object? rows = null;
        object? columns = null;
        try
        {
            if (Convert.ToInt32(((dynamic)shape).HasTable) == 0) return null;
            table = ((dynamic)shape).Table;
            rows = ((dynamic)table).Rows;
            columns = ((dynamic)table).Columns;
            var rowCount = Convert.ToInt32(((dynamic)rows).Count);
            var columnCount = Convert.ToInt32(((dynamic)columns).Count);
            var cells = new List<object>();
            for (var row = 1; row <= rowCount && cells.Count < 50; row++)
            {
                for (var column = 1; column <= columnCount && cells.Count < 50; column++)
                {
                    object? cell = null;
                    object? cellShape = null;
                    try
                    {
                        cell = ((dynamic)table).Cell(row, column);
                        cellShape = ((dynamic)cell).Shape;
                        cells.Add(new { row, column, text = ReadShapeText(cellShape) });
                    }
                    finally
                    {
                        ComInterop.Release(cellShape);
                        ComInterop.Release(cell);
                    }
                }
            }
            return new { rows = rowCount, columns = columnCount, cells };
        }
        catch
        {
            return null;
        }
        finally
        {
            ComInterop.Release(columns);
            ComInterop.Release(rows);
            ComInterop.Release(table);
        }
    }

    private static object? ReadChart(object shape)
    {
        object? chart = null;
        object? title = null;
        object? series = null;
        try
        {
            if (Convert.ToInt32(((dynamic)shape).HasChart) == 0) return null;
            chart = ((dynamic)shape).Chart;
            dynamic chartApi = chart;
            var hasTitle = Safe(() => Convert.ToBoolean(chartApi.HasTitle), false);
            if (hasTitle) title = chartApi.ChartTitle;
            series = Safe<object?>(() => chartApi.SeriesCollection(), null);
            return new
            {
                chartType = Safe(() => Convert.ToInt32(chartApi.ChartType), 0),
                hasTitle,
                title = title is null ? string.Empty : Safe(() => Convert.ToString(((dynamic)title).Text) ?? string.Empty, string.Empty),
                seriesCount = series is null ? 0 : Safe(() => Convert.ToInt32(((dynamic)series).Count), 0),
            };
        }
        catch
        {
            return null;
        }
        finally
        {
            ComInterop.Release(series);
            ComInterop.Release(title);
            ComInterop.Release(chart);
        }
    }

    private static object ReadPicture(object shape)
    {
        object? format = null;
        try
        {
            format = ((dynamic)shape).PictureFormat;
            dynamic api = format;
            return new
            {
                cropLeft = Safe(() => Convert.ToDouble(api.CropLeft), 0d),
                cropRight = Safe(() => Convert.ToDouble(api.CropRight), 0d),
                cropTop = Safe(() => Convert.ToDouble(api.CropTop), 0d),
                cropBottom = Safe(() => Convert.ToDouble(api.CropBottom), 0d),
                lockAspectRatio = Safe(() => Convert.ToInt32(((dynamic)shape).LockAspectRatio), 0),
            };
        }
        finally
        {
            ComInterop.Release(format);
        }
    }

    private static int SlideCount(PresentationActionContext context)
    {
        object? slides = null;
        try
        {
            slides = context.Presentation.Slides;
            return Convert.ToInt32(((dynamic)slides).Count);
        }
        finally
        {
            ComInterop.Release(slides);
        }
    }

    private static string ShapeTypeName(int type) => type switch
    {
        1 => "autoShape",
        3 => "chart",
        6 => "group",
        7 => "embeddedObject",
        13 => "picture",
        14 => "placeholder",
        17 => "textBox",
        19 => "table",
        _ => $"type-{type}",
    };

    private static T Safe<T>(Func<T> value, T fallback)
    {
        try { return value(); }
        catch { return fallback; }
    }

    private sealed record SlideSnapshot(int Index, string Name, string Layout, List<ShapeSnapshot> Shapes, List<object> Overlaps);

    internal sealed record ThemeInspection(object Theme, string ProgId);

    private sealed record ShapeSnapshot(
        int Id,
        string Name,
        int Type,
        string TypeName,
        int ZOrder,
        double Left,
        double Top,
        double Width,
        double Height,
        double Rotation,
        string Text,
        object TextBounds,
        bool TextOverflow,
        bool OutOfBounds,
        object? Table,
        object? Chart,
        object? Picture);
}
