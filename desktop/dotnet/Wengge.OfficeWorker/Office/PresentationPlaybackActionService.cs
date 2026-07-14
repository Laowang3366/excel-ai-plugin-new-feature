using System.Text.Json;
using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Protocol;
using O = Microsoft.Office.Core;
using P = Microsoft.Office.Interop.PowerPoint;

namespace Wengge.OfficeWorker.Office;

internal sealed class PresentationPlaybackActionService(OfficeApplicationProvider applications)
{
    private static readonly HashSet<string> Operations =
    [
        "inspectAnimations",
        "configureAnimations",
        "configureSlideShow",
        "setSpeakerNotes",
        "inspectSpeakerNotes",
        "exportHandouts",
    ];

    public static bool Supports(string operation) => Operations.Contains(operation);

    public object Execute(OfficeActionRequest request)
    {
        if (!Supports(request.Operation))
            throw new OfficeWorkerException("unsupported_operation", $"不支持的演示文稿播放操作: {request.Operation}");

        if (request.Operation == "setSpeakerNotes") return SetSpeakerNotesAndClose(request);
        if (request.Operation == "inspectSpeakerNotes") return InspectSpeakerNotesAndClose(request);
        using var context = new PresentationActionContext(applications, request);
        return request.Operation switch
        {
            "inspectAnimations" => Done(context, request, "已检查动画和切换效果", InspectAnimations(context, request)),
            "configureAnimations" => SaveAndDone(context, request, "已配置动画", ConfigureAnimations(context, request), "animation"),
            "configureSlideShow" => SaveAndDone(context, request, "已配置放映", ConfigureSlideShow(context, request), "slide-show"),
            "exportHandouts" => ExportHandouts(context, request),
            _ => throw new OfficeWorkerException("unsupported_operation", $"不支持的演示文稿播放操作: {request.Operation}"),
        };
    }

    private static object Done(PresentationActionContext context, OfficeActionRequest request, string summary, object data) =>
        OfficeActionResults.Done(request, "com", summary, OfficeActionResults.WithProgId(data, context.ProgId));

    private static object SaveAndDone(PresentationActionContext context, OfficeActionRequest request, string summary, object data, string kind)
    {
        context.Save(request);
        return OfficeActionResults.Done(request, "com", summary, OfficeActionResults.WithProgId(data, context.ProgId),
            [new OfficeChange(kind, request.Target ?? "presentation", summary)]);
    }

    private object SetSpeakerNotesAndClose(OfficeActionRequest request)
    {
        if (UseWpsPackagePath(request)) return SetWpsSpeakerNotesPackage(request);
        SpeakerNotesPlan plan;
        using (var context = new PresentationActionContext(applications, request))
        {
            plan = SetSpeakerNotes(context, request);
            if (!plan.PackageOnly) context.Save(request);
        }
        var outputPath = string.IsNullOrWhiteSpace(request.OutputPath) ? request.FilePath : request.OutputPath;
        if (plan.PackageOnly && !SamePath(request.FilePath, outputPath))
        {
            if (string.IsNullOrWhiteSpace(request.FilePath) || string.IsNullOrWhiteSpace(outputPath))
                throw new OfficeWorkerException("invalid_params", "WPS 备注包级回写缺少源文件或输出路径");
            var fullOutputPath = Path.GetFullPath(outputPath);
            Directory.CreateDirectory(Path.GetDirectoryName(fullOutputPath) ?? Environment.CurrentDirectory);
            File.Copy(Path.GetFullPath(request.FilePath), fullOutputPath, true);
        }
        var packageFallback = plan.PackageUpdates.Count > 0
            && PresentationNotesPackageUpdater.Apply(outputPath, plan.PackageUpdates);
        return OfficeActionResults.Done(request, "com", "已写入演讲者备注", new
        {
            updatedSlides = plan.UpdatedSlides,
            notesPackageFallback = packageFallback,
            progId = plan.ProgId,
        }, [new OfficeChange("speaker-notes", request.Target ?? "presentation", "已写入演讲者备注")]);
    }

    private static object SetWpsSpeakerNotesPackage(OfficeActionRequest request)
    {
        var updates = RequestedPackageUpdates(request);
        var outputPath = PreparePackageOutput(request);
        PresentationNotesPackageUpdater.Apply(outputPath, updates);
        return OfficeActionResults.Done(request, "openxml", "已通过包级兼容路径写入 WPS 演讲者备注", new
        {
            updatedSlides = updates.Select(update => update.SlideIndex).ToArray(),
            notesPackageFallback = true,
            progId = "Wpp.Application",
        }, [new OfficeChange("speaker-notes", request.Target ?? "presentation", "已写入演讲者备注")]);
    }

    private object InspectSpeakerNotesAndClose(OfficeActionRequest request)
    {
        if (UseWpsPackagePath(request))
        {
            var packageNotes = PackageNoteSnapshots(request);
            return OfficeActionResults.Done(request, "openxml", "已通过包级兼容路径检查 WPS 演讲者备注",
                SpeakerNotesResult(packageNotes, "Wpp.Application"));
        }
        string progId;
        using (var context = new PresentationActionContext(applications, request))
        {
            progId = context.ProgId;
            if (!OfficeHostRouting.IsWps(progId))
                return Done(context, request, "已检查演讲者备注", InspectSpeakerNotes(context, request, progId));
        }
        var notes = PresentationNotesPackageUpdater.Read(
                request.FilePath,
                request.BoolParam("allSlides", true),
                request.SlideIndex())
            .Select(note => new NoteSnapshot(
                note.SlideIndex,
                note.SlideText,
                note.NotesText,
                !string.IsNullOrWhiteSpace(note.NotesText),
                Correspondence(note.SlideText, note.NotesText)))
            .ToList();
        return OfficeActionResults.Done(request, "com", "已检查演讲者备注", SpeakerNotesResult(notes, progId));
    }

    private static List<NoteSnapshot> PackageNoteSnapshots(OfficeActionRequest request) =>
        PresentationNotesPackageUpdater.Read(
                request.FilePath,
                request.BoolParam("allSlides", true),
                request.SlideIndex())
            .Select(note => new NoteSnapshot(
                note.SlideIndex,
                note.SlideText,
                note.NotesText,
                !string.IsNullOrWhiteSpace(note.NotesText),
                Correspondence(note.SlideText, note.NotesText)))
            .ToList();

    private static List<PresentationNoteUpdate> RequestedPackageUpdates(OfficeActionRequest request)
    {
        var updates = new List<PresentationNoteUpdate>();
        var notesBySlide = request.Param("notesBySlide");
        if (notesBySlide.ValueKind == JsonValueKind.Array)
        {
            foreach (var note in notesBySlide.EnumerateArray())
            {
                var index = Int(note, "slideIndex");
                if (index < 1) throw new OfficeWorkerException("invalid_params", $"备注页序号必须大于 0: {index}");
                updates.Add(new PresentationNoteUpdate(index, String(note, "text"), Bool(note, "append")));
            }
        }
        else
        {
            if (!request.Params.TryGetProperty("text", out _))
                throw new OfficeWorkerException("invalid_params", "setSpeakerNotes 需要 params.text 或 params.notesBySlide");
            updates.Add(new PresentationNoteUpdate(request.SlideIndex(), request.StringParam("text"), request.BoolParam("append")));
        }
        if (updates.Count == 0) throw new OfficeWorkerException("invalid_params", "notesBySlide 至少需要一项备注");
        return updates;
    }

    private static string PreparePackageOutput(OfficeActionRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.FilePath))
            throw new OfficeWorkerException("invalid_params", "WPS 备注包级回写缺少源文件路径");
        var sourcePath = Path.GetFullPath(request.FilePath);
        var outputPath = Path.GetFullPath(string.IsNullOrWhiteSpace(request.OutputPath) ? request.FilePath : request.OutputPath);
        if (!SamePath(sourcePath, outputPath))
        {
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? Environment.CurrentDirectory);
            File.Copy(sourcePath, outputPath, true);
        }
        return outputPath;
    }

    private static bool UseWpsPackagePath(OfficeActionRequest request) =>
        OfficeHostRouting.RequestsWps(request.StringParam("host"))
        && string.IsNullOrWhiteSpace(request.StringParam("instanceId"))
        && !string.IsNullOrWhiteSpace(request.FilePath);

    private static object InspectAnimations(PresentationActionContext context, OfficeActionRequest request)
    {
        var indexes = request.BoolParam("allSlides")
            ? Enumerable.Range(1, SlideCount(context)).ToArray()
            : [request.SlideIndex()];
        var animations = new List<object>();
        foreach (var index in indexes)
        {
            object? slide = null;
            try
            {
                slide = context.GetSlide(index);
                animations.Add(AnimationSnapshot(slide));
            }
            finally
            {
                ComInterop.Release(slide);
            }
        }

        object? settings = null;
        try
        {
            settings = context.Presentation.SlideShowSettings;
            dynamic api = settings;
            return new
            {
                animations,
                slideShow = new
                {
                    showType = Safe(() => Convert.ToInt32(api.ShowType), 0),
                    advanceMode = Safe(() => Convert.ToInt32(api.AdvanceMode), 0),
                    loopUntilStopped = Safe(() => Convert.ToBoolean(api.LoopUntilStopped), false),
                    showWithAnimation = Safe(() => Convert.ToBoolean(api.ShowWithAnimation), false),
                },
            };
        }
        finally
        {
            ComInterop.Release(settings);
        }
    }

    private static object AnimationSnapshot(object slide)
    {
        object? timeline = null;
        object? sequence = null;
        object? transition = null;
        try
        {
            dynamic slideApi = slide;
            timeline = slideApi.TimeLine;
            sequence = ((dynamic)timeline).MainSequence;
            dynamic sequenceApi = sequence;
            var effects = new List<object>();
            for (var effectIndex = 1; effectIndex <= Convert.ToInt32(sequenceApi.Count); effectIndex++)
            {
                object? effect = null;
                object? effectShape = null;
                object? timing = null;
                object? behaviors = null;
                try
                {
                    effect = sequenceApi.Item(effectIndex);
                    dynamic effectApi = effect;
                    effectShape = Safe<object?>(() => effectApi.Shape, null);
                    timing = Safe<object?>(() => effectApi.Timing, null);
                    behaviors = Safe<object?>(() => effectApi.Behaviors, null);
                    var behaviorSnapshots = new List<object>();
                    if (behaviors is not null)
                    {
                        dynamic behaviorsApi = behaviors;
                        for (var behaviorIndex = 1; behaviorIndex <= Safe(() => Convert.ToInt32(behaviorsApi.Count), 0); behaviorIndex++)
                        {
                            object? behavior = null;
                            object? motion = null;
                            try
                            {
                                behavior = behaviorsApi.Item(behaviorIndex);
                                dynamic behaviorApi = behavior;
                                motion = Safe<object?>(() => behaviorApi.MotionEffect, null);
                                behaviorSnapshots.Add(new
                                {
                                    type = Safe(() => Convert.ToInt32(behaviorApi.Type), 0),
                                    byX = motion is null ? 0 : Safe(() => Convert.ToDouble(((dynamic)motion).ByX), 0d),
                                    byY = motion is null ? 0 : Safe(() => Convert.ToDouble(((dynamic)motion).ByY), 0d),
                                });
                            }
                            finally
                            {
                                ComInterop.Release(motion);
                                ComInterop.Release(behavior);
                            }
                        }
                    }

                    effects.Add(new
                    {
                        index = effectIndex,
                        shapeName = effectShape is null ? string.Empty : Safe(() => Convert.ToString(((dynamic)effectShape).Name) ?? string.Empty, string.Empty),
                        effectType = Safe(() => Convert.ToInt32(effectApi.EffectType), 0),
                        exit = Safe(() => Convert.ToBoolean(effectApi.Exit), false),
                        trigger = timing is null ? 0 : Safe(() => Convert.ToInt32(((dynamic)timing).TriggerType), 0),
                        duration = timing is null ? 0 : Safe(() => Convert.ToDouble(((dynamic)timing).Duration), 0d),
                        delay = timing is null ? 0 : Safe(() => Convert.ToDouble(((dynamic)timing).TriggerDelayTime), 0d),
                        repeatCount = timing is null ? 0 : Safe(() => Convert.ToDouble(((dynamic)timing).RepeatCount), 0d),
                        behaviors = behaviorSnapshots,
                    });
                }
                finally
                {
                    ComInterop.Release(behaviors);
                    ComInterop.Release(timing);
                    ComInterop.Release(effectShape);
                    ComInterop.Release(effect);
                }
            }

            transition = slideApi.SlideShowTransition;
            dynamic transitionApi = transition;
            return new
            {
                slideIndex = Convert.ToInt32(slideApi.SlideIndex),
                effects,
                transition = new
                {
                    entryEffect = Safe(() => Convert.ToInt32(transitionApi.EntryEffect), 0),
                    advanceOnClick = Safe(() => Convert.ToBoolean(transitionApi.AdvanceOnClick), false),
                    advanceOnTime = Safe(() => Convert.ToBoolean(transitionApi.AdvanceOnTime), false),
                    advanceTime = Safe(() => Convert.ToDouble(transitionApi.AdvanceTime), 0d),
                    duration = Safe(() => Convert.ToDouble(transitionApi.Duration), 0d),
                },
            };
        }
        finally
        {
            ComInterop.Release(transition);
            ComInterop.Release(sequence);
            ComInterop.Release(timeline);
        }
    }

    private static object ConfigureAnimations(PresentationActionContext context, OfficeActionRequest request)
    {
        object? slide = null;
        object? timeline = null;
        object? sequence = null;
        object? shapes = null;
        try
        {
            slide = context.GetSlide(request.SlideIndex());
            timeline = ((dynamic)slide).TimeLine;
            sequence = ((dynamic)timeline).MainSequence;
            dynamic sequenceApi = sequence;
            if (request.BoolParam("clearExisting"))
            {
                while (Convert.ToInt32(sequenceApi.Count) > 0)
                {
                    object? effect = null;
                    try { effect = sequenceApi.Item(1); ((dynamic)effect).Delete(); }
                    finally { ComInterop.Release(effect); }
                }
            }

            var effectsParam = request.Param("effects");
            var rules = effectsParam.ValueKind == JsonValueKind.Array ? effectsParam.EnumerateArray().ToArray() : [request.Params];
            shapes = ((dynamic)slide).Shapes;
            dynamic shapesApi = shapes;
            var animated = new List<object>();
            foreach (var rule in rules)
            {
                var category = String(rule, "category", "entrance");
                var names = StringArray(rule, "shapeNames").ToHashSet(StringComparer.OrdinalIgnoreCase);
                var singleName = String(rule, "shapeName");
                if (singleName.Length > 0) names.Add(singleName);
                for (var shapeIndex = 1; shapeIndex <= Convert.ToInt32(shapesApi.Count); shapeIndex++)
                {
                    object? shape = null;
                    object? effect = null;
                    object? timing = null;
                    object? behaviors = null;
                    object? behavior = null;
                    object? motion = null;
                    try
                    {
                        shape = shapesApi.Item(shapeIndex);
                        var shapeName = Safe(() => Convert.ToString(((dynamic)shape).Name) ?? string.Empty, string.Empty);
                        if (names.Count > 0 && !names.Contains(shapeName)) continue;
                        var effectId = AnimationEffect(String(rule, "effect", "fade"), category);
                        var trigger = AnimationTrigger(String(rule, "trigger"));
                        var order = Int(rule, "order");
                        effect = order > 0
                            ? sequenceApi.AddEffect(shape, effectId, 0, trigger, order)
                            : sequenceApi.AddEffect(shape, effectId, 0, trigger);
                        dynamic effectApi = effect;
                        if (category == "exit") effectApi.Exit = -1;
                        if (category == "path")
                        {
                            behaviors = effectApi.Behaviors;
                            behavior = ((dynamic)behaviors).Add(1);
                            motion = ((dynamic)behavior).MotionEffect;
                            ((dynamic)motion).ByX = Double(rule, "pathX", 0.2);
                            ((dynamic)motion).ByY = Double(rule, "pathY", 0);
                        }
                        timing = effectApi.Timing;
                        if (HasNumber(rule, "duration")) ((dynamic)timing).Duration = Double(rule, "duration");
                        if (HasNumber(rule, "delay")) ((dynamic)timing).TriggerDelayTime = Double(rule, "delay");
                        if (HasNumber(rule, "repeatCount")) ((dynamic)timing).RepeatCount = Double(rule, "repeatCount");
                        animated.Add(new { shapeName, category, effect = String(rule, "effect", "fade") });
                    }
                    finally
                    {
                        ComInterop.Release(motion);
                        ComInterop.Release(behavior);
                        ComInterop.Release(behaviors);
                        ComInterop.Release(timing);
                        ComInterop.Release(effect);
                        ComInterop.Release(shape);
                    }
                }
            }
            return new { animated, snapshot = AnimationSnapshot(slide) };
        }
        finally
        {
            ComInterop.Release(shapes);
            ComInterop.Release(sequence);
            ComInterop.Release(timeline);
            ComInterop.Release(slide);
        }
    }

    private static object ConfigureSlideShow(PresentationActionContext context, OfficeActionRequest request)
    {
        object? settings = null;
        try
        {
            settings = context.Presentation.SlideShowSettings;
            dynamic settingsApi = settings;
            settingsApi.ShowType = request.StringParam("showType") switch { "window" => 2, "kiosk" => 3, _ => 1 };
            settingsApi.AdvanceMode = request.BoolParam("autoPlay") || request.BoolParam("useSlideTimings") ? 2 : 1;
            settingsApi.LoopUntilStopped = request.BoolParam("loop") ? -1 : 0;
            settingsApi.ShowWithAnimation = request.BoolParam("showWithAnimation", true) ? -1 : 0;
            var indexes = request.BoolParam("allSlides", true)
                ? Enumerable.Range(1, SlideCount(context)).ToArray()
                : [request.SlideIndex()];
            var entryEffect = request.StringParam("transition") switch
            {
                "cut" => 257,
                "dissolve" => 1537,
                "wipe" => 2817,
                "none" => 0,
                _ => 1793,
            };
            foreach (var index in indexes)
            {
                object? slide = null;
                object? transition = null;
                try
                {
                    slide = context.GetSlide(index);
                    transition = ((dynamic)slide).SlideShowTransition;
                    dynamic api = transition;
                    api.EntryEffect = entryEffect;
                    api.AdvanceOnClick = request.BoolParam("advanceOnClick", true) ? -1 : 0;
                    api.AdvanceOnTime = request.BoolParam("autoPlay") ? -1 : 0;
                    if (request.DoubleParam("advanceAfter") > 0) api.AdvanceTime = request.DoubleParam("advanceAfter");
                    if (request.DoubleParam("transitionDuration") > 0) api.Duration = request.DoubleParam("transitionDuration");
                }
                finally
                {
                    ComInterop.Release(transition);
                    ComInterop.Release(slide);
                }
            }
            return new
            {
                slideShow = new
                {
                    showType = Convert.ToInt32(settingsApi.ShowType),
                    advanceMode = Convert.ToInt32(settingsApi.AdvanceMode),
                    loop = Convert.ToBoolean(settingsApi.LoopUntilStopped),
                    slides = indexes.Length,
                },
            };
        }
        finally
        {
            ComInterop.Release(settings);
        }
    }

    private static SpeakerNotesPlan SetSpeakerNotes(PresentationActionContext context, OfficeActionRequest request)
    {
        var written = new List<int>();
        var packageUpdates = new List<PresentationNoteUpdate>();
        var forcePackageUpdate = OfficeHostRouting.IsWps(context.ProgId);
        var notesBySlide = request.Param("notesBySlide");
        if (notesBySlide.ValueKind == JsonValueKind.Array)
        {
            var slideCount = SlideCount(context);
            foreach (var note in notesBySlide.EnumerateArray())
            {
                var index = Int(note, "slideIndex");
                if (index < 1 || index > slideCount) continue;
                if (forcePackageUpdate) packageUpdates.Add(new PresentationNoteUpdate(index, String(note, "text"), Bool(note, "append")));
                else WriteNotes(context, index, String(note, "text"), Bool(note, "append"), false, packageUpdates);
                written.Add(index);
            }
        }
        else
        {
            if (!request.Params.TryGetProperty("text", out _))
                throw new OfficeWorkerException("invalid_params", "setSpeakerNotes 需要 params.text 或 params.notesBySlide");
            var slideIndex = request.SlideIndex();
            var slideCount = SlideCount(context);
            if (slideIndex < 1 || slideIndex > slideCount) throw new OfficeWorkerException("slide_not_found", $"幻灯片序号超出范围: {slideIndex}");
            if (forcePackageUpdate) packageUpdates.Add(new PresentationNoteUpdate(slideIndex, request.StringParam("text"), request.BoolParam("append")));
            else WriteNotes(context, slideIndex, request.StringParam("text"), request.BoolParam("append"), false, packageUpdates);
            written.Add(slideIndex);
        }
        return new SpeakerNotesPlan(written, packageUpdates, context.ProgId, forcePackageUpdate);
    }

    private static void WriteNotes(
        PresentationActionContext context,
        int slideIndex,
        string text,
        bool append,
        bool forcePackageUpdate,
        List<PresentationNoteUpdate> packageUpdates)
    {
        object? slide = null;
        try
        {
            slide = context.GetSlide(slideIndex);
            var existing = ReadSlideNotesText(slide);
            var applied = SetSlideNotesText(slide, text, append);
            if (!applied && !forcePackageUpdate)
                throw new OfficeWorkerException("notes_not_supported", $"当前 Office 版本无法通过 COM 写入第 {slideIndex} 页备注");
            if (forcePackageUpdate || !applied)
            {
                var finalText = applied ? ReadSlideNotesText(slide)
                    : append && !string.IsNullOrWhiteSpace(existing) ? existing.Trim() + Environment.NewLine + text
                    : text;
                packageUpdates.Add(new PresentationNoteUpdate(slideIndex, finalText, false));
            }
        }
        finally
        {
            ComInterop.Release(slide);
        }
    }

    private static object InspectSpeakerNotes(PresentationActionContext context, OfficeActionRequest request, string progId)
    {
        var indexes = request.BoolParam("allSlides", true)
            ? Enumerable.Range(1, SlideCount(context)).ToArray()
            : [request.SlideIndex()];
        var notes = new List<NoteSnapshot>();
        foreach (var index in indexes)
        {
            object? slide = null;
            try
            {
                slide = context.GetSlide(index);
                var slideText = ReadVisibleText(slide);
                var notesText = ReadSlideNotesText(slide);
                notes.Add(new NoteSnapshot(index, slideText, notesText, !string.IsNullOrWhiteSpace(notesText), Correspondence(slideText, notesText)));
            }
            finally
            {
                ComInterop.Release(slide);
            }
        }
        return SpeakerNotesResult(notes, progId);
    }

    private static object SpeakerNotesResult(List<NoteSnapshot> notes, string progId) => new
        {
            progId,
            notes,
            summary = new
            {
                slideCount = notes.Count,
                missingNotes = notes.Count(note => !note.HasNotes),
                lowCorrespondence = notes.Count(note => note.HasNotes && note.CorrespondenceScore < 0.2),
            },
        };

    private static object ExportHandouts(PresentationActionContext context, OfficeActionRequest request)
    {
        var outputPath = Path.GetFullPath(request.OutputPath ?? Path.Combine(
            Path.GetDirectoryName(request.FilePath) ?? Environment.CurrentDirectory,
            "handouts.pdf"));
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? Environment.CurrentDirectory);
        var outputType = request.BoolParam("includeNotes") || request.StringParam("layout") == "notes"
            ? 5
            : request.StringParam("layout") switch
            {
                "one" => 10,
                "two" => 2,
                "three" => 3,
                "four" => 8,
                "six" => 4,
                "nine" => 9,
                "outline" => 6,
                _ => 3,
            };
        object? printOptions = null;
        object? printRanges = null;
        object? printRange = null;
        try
        {
            printOptions = context.Presentation.PrintOptions;
            ((dynamic)printOptions).OutputType = outputType;
            if (context.ProgId.Equals("PowerPoint.Application", StringComparison.OrdinalIgnoreCase))
            {
                var presentation = (P._Presentation)context.Presentation;
                printRanges = ((P.PrintOptions)printOptions).Ranges;
                printRange = ((P.PrintRanges)printRanges).Add(1, SlideCount(context));
                presentation.ExportAsFixedFormat(
                    outputPath,
                    P.PpFixedFormatType.ppFixedFormatTypePDF,
                    P.PpFixedFormatIntent.ppFixedFormatIntentPrint,
                    O.MsoTriState.msoFalse,
                    P.PpPrintHandoutOrder.ppPrintHandoutVerticalFirst,
                    (P.PpPrintOutputType)outputType,
                    O.MsoTriState.msoFalse,
                    (P.PrintRange)printRange,
                    P.PpPrintRangeType.ppPrintSlideRange,
                    string.Empty,
                    true, true, true, true, false,
                    Type.Missing);
            }
            else
            {
                context.Presentation.ExportAsFixedFormat(
                    outputPath, 2, 2, 0, 1, outputType, 0, Type.Missing, 1, string.Empty,
                    true, true, true, true, false, Type.Missing);
            }
        }
        catch (Exception exception)
        {
            throw new OfficeWorkerException("export_failed", $"当前 PowerPoint/WPS 版本无法导出指定讲义版式: {exception.GetBaseException().Message}", null, exception);
        }
        finally
        {
            ComInterop.Release(printRange);
            ComInterop.Release(printRanges);
            ComInterop.Release(printOptions);
        }
        if (!File.Exists(outputPath)) throw new OfficeWorkerException("export_failed", "PowerPoint 讲义 PDF 未生成");
        return OfficeActionResults.Done(request, "com", "已导出讲义 PDF",
            new { progId = context.ProgId, outputType, includeNotes = outputType == 5 },
            [new OfficeChange("export", outputPath, "已导出带备注或讲义版 PDF")], outputPath);
    }

    private static bool SetSlideNotesText(object slide, string text, bool append)
    {
        object? notesPage = null;
        object? shapes = null;
        try
        {
            notesPage = ((dynamic)slide).NotesPage;
            shapes = ((dynamic)notesPage).Shapes;
            dynamic shapesApi = shapes;
            for (var index = 1; index <= Convert.ToInt32(shapesApi.Count); index++)
            {
                object? shape = null;
                object? placeholder = null;
                object? frame = null;
                object? range = null;
                try
                {
                    shape = shapesApi.Item(index);
                    placeholder = ((dynamic)shape).PlaceholderFormat;
                    if (Convert.ToInt32(((dynamic)placeholder).Type) != 2) continue;
                    frame = ((dynamic)shape).TextFrame;
                    range = ((dynamic)frame).TextRange;
                    var existing = Convert.ToString(((dynamic)range).Text) ?? string.Empty;
                    ((dynamic)range).Text = append && !string.IsNullOrWhiteSpace(existing)
                        ? existing.Trim() + Environment.NewLine + text
                        : text;
                    return true;
                }
                catch
                {
                    // Continue because notes pages also contain non-text placeholders.
                }
                finally
                {
                    ComInterop.Release(range);
                    ComInterop.Release(frame);
                    ComInterop.Release(placeholder);
                    ComInterop.Release(shape);
                }
            }
            return false;
        }
        finally
        {
            ComInterop.Release(shapes);
            ComInterop.Release(notesPage);
        }
    }

    private sealed record SpeakerNotesPlan(List<int> UpdatedSlides, List<PresentationNoteUpdate> PackageUpdates, string ProgId, bool PackageOnly);

    private static bool SamePath(string? first, string? second) =>
        !string.IsNullOrWhiteSpace(first) && !string.IsNullOrWhiteSpace(second)
        && string.Equals(Path.GetFullPath(first), Path.GetFullPath(second), StringComparison.OrdinalIgnoreCase);

    private static string ReadSlideNotesText(object slide)
    {
        object? notesPage = null;
        object? shapes = null;
        try
        {
            notesPage = ((dynamic)slide).NotesPage;
            shapes = ((dynamic)notesPage).Shapes;
            dynamic shapesApi = shapes;
            for (var index = 1; index <= Convert.ToInt32(shapesApi.Count); index++)
            {
                object? shape = null;
                object? placeholder = null;
                object? frame = null;
                object? range = null;
                try
                {
                    shape = shapesApi.Item(index);
                    placeholder = ((dynamic)shape).PlaceholderFormat;
                    if (Convert.ToInt32(((dynamic)placeholder).Type) != 2) continue;
                    frame = ((dynamic)shape).TextFrame;
                    range = ((dynamic)frame).TextRange;
                    return Convert.ToString(((dynamic)range).Text) ?? string.Empty;
                }
                catch
                {
                    // Continue because notes pages also contain non-text placeholders.
                }
                finally
                {
                    ComInterop.Release(range);
                    ComInterop.Release(frame);
                    ComInterop.Release(placeholder);
                    ComInterop.Release(shape);
                }
            }
            return string.Empty;
        }
        catch
        {
            return string.Empty;
        }
        finally
        {
            ComInterop.Release(shapes);
            ComInterop.Release(notesPage);
        }
    }

    private static string ReadVisibleText(object slide)
    {
        object? shapes = null;
        try
        {
            shapes = ((dynamic)slide).Shapes;
            dynamic shapesApi = shapes;
            var parts = new List<string>();
            for (var index = 1; index <= Convert.ToInt32(shapesApi.Count); index++)
            {
                object? shape = null;
                object? frame = null;
                object? range = null;
                try
                {
                    shape = shapesApi.Item(index);
                    if (Convert.ToInt32(((dynamic)shape).HasTextFrame) == 0) continue;
                    frame = ((dynamic)shape).TextFrame;
                    if (Convert.ToInt32(((dynamic)frame).HasText) == 0) continue;
                    range = ((dynamic)frame).TextRange;
                    var text = Convert.ToString(((dynamic)range).Text);
                    if (!string.IsNullOrWhiteSpace(text)) parts.Add(text);
                }
                catch
                {
                    // Ignore shapes that do not expose a text frame.
                }
                finally
                {
                    ComInterop.Release(range);
                    ComInterop.Release(frame);
                    ComInterop.Release(shape);
                }
            }
            return string.Join(Environment.NewLine, parts);
        }
        finally
        {
            ComInterop.Release(shapes);
        }
    }

    private static double Correspondence(string slideText, string notesText)
    {
        if (string.IsNullOrWhiteSpace(slideText) || string.IsNullOrWhiteSpace(notesText)) return 0;
        var slideTerms = slideText.ToLowerInvariant().Where(char.IsLetterOrDigit).ToHashSet();
        var noteTerms = notesText.ToLowerInvariant().Where(char.IsLetterOrDigit).ToHashSet();
        if (slideTerms.Count == 0 || noteTerms.Count == 0) return 0;
        return Math.Round(noteTerms.Count(slideTerms.Contains) / (double)Math.Min(slideTerms.Count, noteTerms.Count), 3);
    }

    private static int SlideCount(PresentationActionContext context)
    {
        object? slides = null;
        try { slides = context.Presentation.Slides; return Convert.ToInt32(((dynamic)slides).Count); }
        finally { ComInterop.Release(slides); }
    }

    private static int AnimationEffect(string name, string category) => category == "emphasis"
        ? name switch { "spin" => 61, "transparency" => 62, _ => 59 }
        : name switch { "appear" => 1, "fly" => 2, "dissolve" => 9, "wipe" => 22, "zoom" => 23, _ => 10 };

    private static int AnimationTrigger(string value) => value switch { "withPrevious" => 2, "afterPrevious" => 3, _ => 1 };

    private static string String(JsonElement value, string name, string fallback = "") =>
        value.ValueKind == JsonValueKind.Object && value.TryGetProperty(name, out var property) && property.ValueKind == JsonValueKind.String
            ? property.GetString() ?? fallback
            : fallback;

    private static bool Bool(JsonElement value, string name, bool fallback = false) =>
        value.ValueKind == JsonValueKind.Object && value.TryGetProperty(name, out var property) && property.ValueKind is JsonValueKind.True or JsonValueKind.False
            ? property.GetBoolean()
            : fallback;

    private static int Int(JsonElement value, string name, int fallback = 0) =>
        value.ValueKind == JsonValueKind.Object && value.TryGetProperty(name, out var property) && property.TryGetInt32(out var result) ? result : fallback;

    private static double Double(JsonElement value, string name, double fallback = 0) =>
        value.ValueKind == JsonValueKind.Object && value.TryGetProperty(name, out var property) && property.TryGetDouble(out var result) ? result : fallback;

    private static bool HasNumber(JsonElement value, string name) =>
        value.ValueKind == JsonValueKind.Object && value.TryGetProperty(name, out var property) && property.ValueKind == JsonValueKind.Number;

    private static IEnumerable<string> StringArray(JsonElement value, string name)
    {
        if (value.ValueKind != JsonValueKind.Object || !value.TryGetProperty(name, out var property) || property.ValueKind != JsonValueKind.Array) yield break;
        foreach (var item in property.EnumerateArray())
            if (item.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(item.GetString())) yield return item.GetString()!;
    }

    private static T Safe<T>(Func<T> value, T fallback)
    {
        try { return value(); }
        catch { return fallback; }
    }

    private sealed record NoteSnapshot(int SlideIndex, string SlideText, string NotesText, bool HasNotes, double CorrespondenceScore);
}
