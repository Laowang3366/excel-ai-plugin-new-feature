using System.Text.Json;
using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Office;

internal sealed class ExcelCrossOfficeActionService(OfficeApplicationProvider applications)
{
    public static bool Supports(string operation) => operation is "exportRangeToWord" or "exportRangeToPresentation" or "buildReportPackage";

    public object Execute(OfficeActionRequest request)
    {
        return request.Operation switch
        {
            "exportRangeToWord" => ExportToWord(request),
            "exportRangeToPresentation" => ExportToPresentation(request),
            "buildReportPackage" => BuildReport(request),
            _ => throw new OfficeWorkerException("unsupported_operation", $"不支持的跨软件操作: {request.Operation}"),
        };
    }

    private object ExportToWord(OfficeActionRequest request)
    {
        using var excel = new ExcelActionContext(applications, OfficeHostRouting.SourceRequest(request, "excel", request.FilePath));
        var output = OutputPath(request, request.StringParam("wordOutputPath", request.OutputPath ?? "report.docx"), ".docx");
        var updateExisting = request.BoolParam("updateExisting");
        var targetInstanceId = TargetInstanceId(request, "word");
        var linked = request.BoolParam("linked");
        var linkId = request.StringParam("linkId", Guid.NewGuid().ToString("N"));
        if (updateExisting && request.StringParam("linkId").Length == 0)
            throw new OfficeWorkerException("invalid_params", "exportRangeToWord 增量更新需要 params.linkId");
        if (targetInstanceId.Length == 0) PrepareOutput(output, updateExisting, request.BoolParam("overwrite"));
        using var word = TargetApplication.Open(applications, "word", request.StringParam("wordHost"), targetInstanceId, output);
        dynamic app = word.Application;
        object? documents = null;
        object? document = null;
        object? range = null;
        object? existingShape = null;
        object? pastedShape = null;
        IReadOnlyList<WordBookmarkSnapshot> preservedBookmarks = [];
        IReadOnlyList<string> restoredBookmarks = [];
        var openedDocument = false;
        try
        {
            if (word.Document is not null) document = word.Document;
            else
            {
                documents = app.Documents;
                dynamic documentsApi = documents;
                document = File.Exists(output) && updateExisting ? documentsApi.Open(output) : documentsApi.Add();
                openedDocument = true;
            }
            dynamic documentApi = document;
            if (updateExisting) preservedBookmarks = CaptureWordBookmarks(documentApi);
            existingShape = FindWordManagedShape(documentApi, linkId);
            if (updateExisting && existingShape is null && !request.BoolParam("allowMissingManaged"))
                throw new OfficeWorkerException("linked_object_not_found", $"找不到 Word 受管链接对象: {linkId}");
            if (existingShape is not null)
            {
                object? existingRange = null;
                try
                {
                    existingRange = ((dynamic)existingShape).Range;
                    var start = Convert.ToInt32(((dynamic)existingRange).Start);
                    ((dynamic)existingShape).Delete();
                    range = documentApi.Range(start, start);
                }
                finally { ComInterop.Release(existingRange); }
            }
            else
            {
                var end = Math.Max(0, Convert.ToInt32(documentApi.Content.End) - 1);
                range = documentApi.Range(end, end);
            }
            dynamic rangeApi = range;
            var title = request.StringParam("title");
            if (title.Length > 0 && existingShape is null)
            {
                rangeApi.InsertAfter(title + Environment.NewLine);
                rangeApi.Collapse(0);
            }
            var insertionStart = Convert.ToInt32(rangeApi.Start);
            CopyExcelContent(excel, request, asPicture: !linked && request.BoolParam("asPicture"));
            if (linked) rangeApi.PasteSpecial(0, true, 0, false, 0);
            else if (request.BoolParam("asPicture")) rangeApi.PasteSpecial(0, false, 0, false, 9);
            else rangeApi.PasteExcelTable(false, false, false);
            if (linked)
            {
                pastedShape = FindWordInlineShapeAt(documentApi, insertionStart)
                    ?? throw new OfficeWorkerException("linked_object_not_created", "Word 未创建 Excel 链接对象");
                var metadata = JsonSerializer.Serialize(new
                {
                    version = 1,
                    linkId,
                    source = Path.GetFullPath(request.FilePath ?? string.Empty),
                    sourceType = request.StringParam("sourceType", "range"),
                    sourceName = request.StringParam("chartName", request.StringParam("sourceName")),
                    range = request.ExcelTarget().Address,
                    managed = true,
                });
                try { ((dynamic)pastedShape).AlternativeText = "WENGGE_MANIFEST:" + metadata; } catch { }
                SetWordManagedIds((object)document, MergeIds(ReadWordManagedIds((object)document), linkId));
            }
            if (updateExisting) restoredBookmarks = RestoreWordBookmarks(documentApi, preservedBookmarks);
            if (word.Document is not null || File.Exists(output) && updateExisting) documentApi.Save(); else documentApi.SaveAs2(output);
            return Done(request, "cross-office-word", "已将 Excel 区域写入 Word", new
            {
                outputPath = output,
                sourceProgId = excel.ProgId,
                targetProgId = word.ProgId,
                linkedObjects = new[] { new { target = "word", linkId, source = request.FilePath, linked } },
                preservedBookmarks = preservedBookmarks.Count,
                restoredBookmarks,
            }, output);
        }
        finally
        {
            ComInterop.Release(pastedShape);
            ComInterop.Release(existingShape);
            ComInterop.Release(range);
            if (openedDocument) { try { ((dynamic?)document)?.Close(0); } catch { } }
            if (word.Document is null) ComInterop.Release(document);
            ComInterop.Release(documents);
        }
    }

    private object ExportToPresentation(OfficeActionRequest request)
    {
        using var excel = new ExcelActionContext(applications, OfficeHostRouting.SourceRequest(request, "excel", request.FilePath));
        var output = OutputPath(request, request.StringParam("presentationOutputPath", request.OutputPath ?? "report.pptx"), ".pptx");
        var updateExisting = request.BoolParam("updateExisting");
        var targetInstanceId = TargetInstanceId(request, "presentation");
        var linked = request.BoolParam("linked");
        var linkId = request.StringParam("linkId", Guid.NewGuid().ToString("N"));
        if (updateExisting && request.StringParam("linkId").Length == 0)
            throw new OfficeWorkerException("invalid_params", "exportRangeToPresentation 增量更新需要 params.linkId");
        if (targetInstanceId.Length == 0) PrepareOutput(output, updateExisting, request.BoolParam("overwrite"));
        using var powerPoint = TargetApplication.Open(applications, "presentation", request.StringParam("presentationHost"), targetInstanceId, output);
        dynamic app = powerPoint.Application;
        object? presentations = null;
        object? presentation = null;
        object? slides = null;
        object? slide = null;
        object? shapes = null;
        object? pasted = null;
        object? shape = null;
        object? existingShape = null;
        var openedPresentation = false;
        try
        {
            if (powerPoint.Document is not null) presentation = powerPoint.Document;
            else
            {
                presentations = app.Presentations; dynamic presentationsApi = presentations;
                presentation = File.Exists(output) && updateExisting ? presentationsApi.Open(output) : presentationsApi.Add();
                openedPresentation = true;
            }
            dynamic presentationApi = presentation;
            slides = presentationApi.Slides; dynamic slidesApi = slides;
            var managedShape = FindPresentationManagedShape((object)presentation, linkId);
            slide = managedShape.Slide;
            existingShape = managedShape.Shape;
            var createdSlide = slide is null;
            if (updateExisting && existingShape is null && !request.BoolParam("allowMissingManaged"))
                throw new OfficeWorkerException("linked_object_not_found", $"找不到 PowerPoint 受管链接对象: {linkId}");
            slide ??= slidesApi.Add(Convert.ToInt32(slidesApi.Count) + 1, 12);
            dynamic slideApi = slide;
            shapes = slideApi.Shapes; dynamic shapesApi = shapes;
            var geometry = existingShape is null ? null : ShapeGeometry.Read(existingShape);
            var existingName = existingShape is null ? string.Empty : Convert.ToString(((dynamic)existingShape).Name) ?? string.Empty;
            if (existingShape is not null) ((dynamic)existingShape).Delete();
            var title = request.StringParam("title");
            if (title.Length > 0) SetPresentationTitle(presentationApi, slideApi, shapesApi, linkId, title);
            CopyExcelContent(excel, request, asPicture: !linked);
            pasted = linked ? shapesApi.PasteSpecial(10, 0, string.Empty, 0, string.Empty, -1) : shapesApi.PasteSpecial(2);
            dynamic pastedApi = pasted;
            shape = pastedApi.Item(1);
            dynamic shapeApi = shape;
            if (existingName.Length > 0) shapeApi.Name = existingName;
            if (geometry is not null)
            {
                geometry.Apply(shapeApi);
            }
            else
            {
                shapeApi.Left = request.DoubleParam("left", 40);
                shapeApi.Top = request.DoubleParam("top", title.Length > 0 ? 80 : 40);
                if (request.DoubleParam("width") > 0) shapeApi.Width = request.DoubleParam("width");
            }
            SetTag(shapeApi, "WENGGE_LINK_ID", linkId);
            SetTag(shapeApi, "WENGGE_SOURCE_PATH", Path.GetFullPath(request.FilePath ?? string.Empty));
            SetTag(shapeApi, "WENGGE_SOURCE_TYPE", request.StringParam("sourceType", "range"));
            SetTag(shapeApi, "WENGGE_SOURCE_NAME", request.StringParam("chartName", request.StringParam("sourceName")));
            SetTag(shapeApi, "WENGGE_SOURCE_RANGE", request.ExcelTarget().Address);
            SetTag(shapeApi, "WENGGE_MANAGED", "true");
            SetTag(slideApi, "WENGGE_LINK_ID", linkId);
            if (createdSlide) SetTag(slideApi, "WENGGE_MANAGED_SLIDE", "true");
            SetTag(presentationApi, "WENGGE_MANAGED_LINK_IDS", JsonSerializer.Serialize(MergeIds(ReadPresentationManagedIds((object)presentation), linkId)));
            SetTag(presentationApi, "WENGGE_MANIFEST_VERSION", "1");
            if (powerPoint.Document is not null || File.Exists(output) && updateExisting) presentationApi.Save(); else presentationApi.SaveAs(output);
            return Done(request, "cross-office-presentation", "已将 Excel 区域写入演示文稿", new
            {
                outputPath = output,
                sourceProgId = excel.ProgId,
                targetProgId = powerPoint.ProgId,
                slideIndex = Convert.ToInt32(slideApi.SlideIndex),
                linkedObjects = new[] { new { target = "presentation", linkId, shapeName = Convert.ToString(shapeApi.Name), source = request.FilePath, linked } },
            }, output);
        }
        finally
        {
            ComInterop.Release(existingShape);
            ComInterop.Release(shape);
            ComInterop.Release(pasted);
            ComInterop.Release(shapes);
            ComInterop.Release(slide);
            ComInterop.Release(slides);
            if (openedPresentation) { try { ((dynamic?)presentation)?.Close(); } catch { } }
            if (powerPoint.Document is null) ComInterop.Release(presentation);
            ComInterop.Release(presentations);
        }
    }

    private object BuildReport(OfficeActionRequest request)
    {
        var outputDirectory = ReportOutputDirectory(request);
        var baseName = request.StringParam("baseName", $"{Path.GetFileNameWithoutExtension(request.FilePath)}-报告");
        var wordPath = request.StringParam("wordOutputPath", Path.Combine(outputDirectory, baseName + ".docx"));
        var presentationPath = request.StringParam("presentationOutputPath", Path.Combine(outputDirectory, baseName + ".pptx"));
        var sections = ReportSections(request);
        var updateExisting = request.BoolParam("updateExisting");
        if (updateExisting)
        {
            var missing = sections.FindIndex(section => section.LinkId.Length == 0);
            if (missing >= 0)
                throw new OfficeWorkerException("invalid_params", $"buildReportPackage 增量更新的第 {missing + 1} 个 section 缺少稳定 linkId");
        }

        var outputs = new List<object>();
        var wordIds = new List<string>();
        var presentationIds = new List<string>();
        for (var index = 0; index < sections.Count; index++)
        {
            var section = sections[index];
            var sectionId = section.LinkId.Length > 0 ? section.LinkId : Guid.NewGuid().ToString("N");
            var wordId = sectionId + "_word";
            var presentationId = sectionId + "_ppt";
            wordIds.Add(wordId);
            presentationIds.Add(presentationId);

            var wordRequest = ReportSectionRequest(request, section, "exportRangeToWord", wordPath, wordId, updateExisting || index > 0, index == 0);
            outputs.Add(ExportToWord(wordRequest));
        }
        for (var index = 0; index < sections.Count; index++)
        {
            var section = sections[index];
            var presentationRequest = ReportSectionRequest(request, section, "exportRangeToPresentation", presentationPath, presentationIds[index], updateExisting || index > 0, index == 0);
            outputs.Add(ExportToPresentation(presentationRequest));
        }
        return Done(request, "report-package", "已生成跨 Office 报告包", new
        {
            outputDirectory,
            wordOutputPath = Path.GetFullPath(wordPath),
            presentationOutputPath = Path.GetFullPath(presentationPath),
            outputs,
            manifest = new { version = 1, word = wordIds, presentation = presentationIds },
        }, outputDirectory);
    }

    internal static List<ReportSection> ReportSections(OfficeActionRequest request)
    {
        var value = request.Param("sections");
        if (value.ValueKind != JsonValueKind.Array || value.GetArrayLength() == 0)
            return [new ReportSection(request.StringParam("linkId"), JsonSerializer.SerializeToElement(new { }))];

        var sections = new List<ReportSection>();
        foreach (var item in value.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.Object)
                throw new OfficeWorkerException("invalid_params", $"buildReportPackage 的第 {sections.Count + 1} 个 section 必须是对象");
            var linkId = item.TryGetProperty("linkId", out var property) && property.ValueKind == JsonValueKind.String
                ? property.GetString()?.Trim() ?? string.Empty
                : string.Empty;
            sections.Add(new ReportSection(linkId, item.Clone()));
        }
        return sections;
    }

    internal static OfficeActionRequest ReportSectionRequest(
        OfficeActionRequest request,
        ReportSection section,
        string operation,
        string outputPath,
        string linkId,
        bool updateExisting,
        bool firstSection)
    {
        var parameters = request.Params.EnumerateObject()
            .Where(property => property.Name is not ("sections" or "outputDirectory" or "baseName"))
            .ToDictionary(property => property.Name, property => property.Value.Clone(), StringComparer.Ordinal);
        foreach (var property in section.Values.EnumerateObject())
            parameters[property.Name] = property.Value.Clone();
        parameters["linkId"] = JsonSerializer.SerializeToElement(linkId);
        parameters["updateExisting"] = JsonSerializer.SerializeToElement(updateExisting);
        parameters["allowMissingManaged"] = JsonSerializer.SerializeToElement(updateExisting);
        parameters["overwrite"] = JsonSerializer.SerializeToElement(firstSection && request.BoolParam("overwrite"));

        var target = request.Target;
        var range = parameters.TryGetValue("range", out var rangeValue) && rangeValue.ValueKind == JsonValueKind.String ? rangeValue.GetString() : null;
        if (!string.IsNullOrWhiteSpace(range))
        {
            var sheetName = parameters.TryGetValue("sheetName", out var sheetValue) && sheetValue.ValueKind == JsonValueKind.String
                ? sheetValue.GetString()
                : request.ExcelTarget().SheetName;
            target = $"range:{sheetName}!{range}";
        }
        return request with
        {
            Operation = operation,
            OutputPath = outputPath,
            Target = target,
            Params = JsonSerializer.SerializeToElement(parameters),
        };
    }

    private static string ReportOutputDirectory(OfficeActionRequest request)
    {
        var configured = request.StringParam("outputDirectory");
        var directory = configured.Length > 0
            ? configured
            : !string.IsNullOrWhiteSpace(request.OutputPath)
                ? Path.HasExtension(request.OutputPath) ? Path.GetDirectoryName(request.OutputPath)! : request.OutputPath
                : Path.GetDirectoryName(Path.GetFullPath(request.FilePath ?? Environment.CurrentDirectory))!;
        var result = Path.GetFullPath(directory);
        Directory.CreateDirectory(result);
        return result;
    }

    private static void CopyExcelContent(ExcelActionContext context, OfficeActionRequest request, bool asPicture)
    {
        var (sheet, range) = context.GetRange(request);
        object? chartObjects = null;
        object? chartObject = null;
        object? chart = null;
        object? chartArea = null;
        try
        {
            if (request.StringParam("sourceType") == "chart")
            {
                var name = request.StringParam("chartName", request.StringParam("sourceName"));
                if (name.Length == 0) throw new OfficeWorkerException("invalid_params", "图表联动需要 params.chartName 或 params.sourceName");
                chartObjects = ((dynamic)sheet).ChartObjects();
                chartObject = ((dynamic)chartObjects).Item(name);
                chart = ((dynamic)chartObject).Chart;
                ((dynamic)sheet).Activate();
                ((dynamic)chartObject).Activate();
                if (asPicture) ((dynamic)chart).CopyPicture();
                else
                {
                    chartArea = ((dynamic)chart).ChartArea;
                    ((dynamic)chartArea).Copy();
                }
            }
            else if (asPicture) ((dynamic)range).CopyPicture(1, 2);
            else ((dynamic)range).Copy();
        }
        finally
        {
            ComInterop.Release(chartArea);
            ComInterop.Release(chart);
            ComInterop.Release(chartObject);
            ComInterop.Release(chartObjects);
            ComInterop.Release(range);
            ComInterop.Release(sheet);
        }
    }

    private static object? FindWordManagedShape(dynamic document, string linkId)
    {
        object? inlineShapes = null;
        var managedIds = ReadWordManagedIds((object)document);
        try
        {
            inlineShapes = document.InlineShapes;
            dynamic shapesApi = inlineShapes;
            for (var index = 1; index <= Convert.ToInt32(shapesApi.Count); index++)
            {
                object? shape = null;
                try
                {
                    shape = shapesApi.Item(index);
                    var alternativeText = Convert.ToString(((dynamic)shape).AlternativeText) ?? string.Empty;
                    var candidateId = ManifestLinkId(alternativeText);
                    if (candidateId.Length == 0 && index <= managedIds.Length) candidateId = managedIds[index - 1];
                    if (candidateId.Equals(linkId, StringComparison.Ordinal))
                    {
                        var result = shape;
                        shape = null;
                        return result;
                    }
                }
                catch { }
                finally { ComInterop.Release(shape); }
            }
            return null;
        }
        finally { ComInterop.Release(inlineShapes); }
    }

    private static IReadOnlyList<WordBookmarkSnapshot> CaptureWordBookmarks(dynamic document)
    {
        object? bookmarks = null;
        var snapshots = new List<WordBookmarkSnapshot>();
        try
        {
            bookmarks = document.Bookmarks;
            dynamic bookmarksApi = bookmarks;
            for (var index = 1; index <= Convert.ToInt32(bookmarksApi.Count); index++)
            {
                object? bookmark = null;
                object? range = null;
                try
                {
                    bookmark = bookmarksApi.Item(index);
                    range = ((dynamic)bookmark).Range;
                    snapshots.Add(new WordBookmarkSnapshot(
                        Convert.ToString(((dynamic)bookmark).Name) ?? string.Empty,
                        Convert.ToInt32(((dynamic)range).Start),
                        Convert.ToInt32(((dynamic)range).End),
                        Convert.ToString(((dynamic)range).Text) ?? string.Empty));
                }
                finally
                {
                    ComInterop.Release(range);
                    ComInterop.Release(bookmark);
                }
            }
            return snapshots.Where(snapshot => snapshot.Name.Length > 0).ToArray();
        }
        finally { ComInterop.Release(bookmarks); }
    }

    private static IReadOnlyList<string> RestoreWordBookmarks(dynamic document, IReadOnlyList<WordBookmarkSnapshot> snapshots)
    {
        if (snapshots.Count == 0) return [];
        object? bookmarks = null;
        object? content = null;
        var restored = new List<string>();
        try
        {
            bookmarks = document.Bookmarks;
            dynamic bookmarksApi = bookmarks;
            content = document.Content;
            var contentEnd = Math.Max(0, Convert.ToInt32(((dynamic)content).End) - 1);
            var contentText = Convert.ToString(((dynamic)content).Text) ?? string.Empty;
            foreach (var snapshot in snapshots)
            {
                if (Convert.ToBoolean(bookmarksApi.Exists(snapshot.Name))) continue;
                var start = Math.Clamp(snapshot.Start, 0, contentEnd);
                var end = Math.Clamp(snapshot.End, start, contentEnd);
                if (snapshot.Text.Length > 0 && !RangeMatches(document, start, end, snapshot.Text))
                {
                    var located = ClosestOccurrence(contentText, snapshot.Text, snapshot.Start);
                    if (located >= 0)
                    {
                        start = located;
                        end = Math.Min(contentEnd, located + snapshot.Text.Length);
                    }
                }
                object? range = null;
                object? bookmark = null;
                try
                {
                    range = document.Range(start, end);
                    bookmark = bookmarksApi.Add(snapshot.Name, range);
                }
                catch (Exception exception)
                {
                    throw new OfficeWorkerException("bookmark_restore_failed", $"增量更新后无法恢复 Word 书签: {snapshot.Name}", new
                    {
                        snapshot.Name,
                        snapshot.Start,
                        snapshot.End,
                        start,
                        end,
                    }, exception);
                }
                finally
                {
                    ComInterop.Release(bookmark);
                    ComInterop.Release(range);
                }
                if (!Convert.ToBoolean(bookmarksApi.Exists(snapshot.Name)))
                    throw new OfficeWorkerException("bookmark_restore_failed", $"增量更新后 Word 书签恢复校验失败: {snapshot.Name}");
                restored.Add(snapshot.Name);
            }
            return restored;
        }
        finally
        {
            ComInterop.Release(content);
            ComInterop.Release(bookmarks);
        }
    }

    private static bool RangeMatches(dynamic document, int start, int end, string expected)
    {
        object? range = null;
        try
        {
            range = document.Range(start, end);
            return string.Equals(Convert.ToString(((dynamic)range).Text) ?? string.Empty, expected, StringComparison.Ordinal);
        }
        catch { return false; }
        finally { ComInterop.Release(range); }
    }

    private static int ClosestOccurrence(string source, string value, int preferredIndex)
    {
        var closest = -1;
        var closestDistance = int.MaxValue;
        for (var index = source.IndexOf(value, StringComparison.Ordinal); index >= 0; index = source.IndexOf(value, index + 1, StringComparison.Ordinal))
        {
            var distance = Math.Abs(index - preferredIndex);
            if (distance >= closestDistance) continue;
            closest = index;
            closestDistance = distance;
        }
        return closest;
    }

    private static string[] ReadWordManagedIds(object document)
    {
        object? variables = null;
        object? item = null;
        try
        {
            variables = ((dynamic)document).Variables;
            item = ((dynamic)variables).Item("WENGGE_MANAGED_LINK_IDS");
            var value = Convert.ToString((object?)((dynamic)item).Value) ?? string.Empty;
            return JsonSerializer.Deserialize<string[]>(value) ?? [];
        }
        catch { return []; }
        finally { ComInterop.Release(item); ComInterop.Release(variables); }
    }

    private static string[] ReadPresentationManagedIds(object presentation)
    {
        var value = ReadTag(presentation, "WENGGE_MANAGED_LINK_IDS");
        try { return JsonSerializer.Deserialize<string[]>(value) ?? []; }
        catch { return []; }
    }

    private static string[] MergeIds(IEnumerable<string> existing, string linkId) =>
        existing.Append(linkId).Where(id => id.Length > 0).Distinct(StringComparer.Ordinal).ToArray();

    private static void SetWordManagedIds(object document, string[] ids)
    {
        object? variables = null;
        object? item = null;
        try
        {
            variables = ((dynamic)document).Variables;
            var value = JsonSerializer.Serialize(ids.Distinct(StringComparer.Ordinal).ToArray());
            try { item = ((dynamic)variables).Item("WENGGE_MANAGED_LINK_IDS"); ((dynamic)item).Value = value; }
            catch { item = ((dynamic)variables).Add("WENGGE_MANAGED_LINK_IDS", value); }
            ComInterop.Release(item); item = null;
            try { item = ((dynamic)variables).Item("WENGGE_MANIFEST_VERSION"); ((dynamic)item).Value = "1"; }
            catch { item = ((dynamic)variables).Add("WENGGE_MANIFEST_VERSION", "1"); }
        }
        finally { ComInterop.Release(item); ComInterop.Release(variables); }
    }

    private static object? FindWordInlineShapeAt(dynamic document, int start)
    {
        object? inlineShapes = null;
        object? match = null;
        var matchStart = int.MaxValue;
        try
        {
            inlineShapes = document.InlineShapes;
            dynamic shapesApi = inlineShapes;
            for (var index = 1; index <= Convert.ToInt32(shapesApi.Count); index++)
            {
                object? shape = null;
                object? range = null;
                try
                {
                    shape = shapesApi.Item(index);
                    range = ((dynamic)shape).Range;
                    var candidateStart = Convert.ToInt32(((dynamic)range).Start);
                    if (candidateStart < start || candidateStart >= matchStart) continue;
                    ComInterop.Release(match);
                    match = shape;
                    shape = null;
                    matchStart = candidateStart;
                }
                catch { }
                finally { ComInterop.Release(range); ComInterop.Release(shape); }
            }
            var result = match;
            match = null;
            return result;
        }
        finally { ComInterop.Release(match); ComInterop.Release(inlineShapes); }
    }

    private static string ManifestLinkId(string alternativeText)
    {
        if (!alternativeText.StartsWith("WENGGE_MANIFEST:", StringComparison.Ordinal)) return string.Empty;
        try
        {
            using var document = JsonDocument.Parse(alternativeText[16..]);
            return document.RootElement.TryGetProperty("linkId", out var value) ? value.GetString() ?? string.Empty : string.Empty;
        }
        catch { return string.Empty; }
    }

    private static (object? Slide, object? Shape) FindPresentationManagedShape(object presentation, string linkId)
    {
        object? slides = null;
        try
        {
            slides = ((dynamic)presentation).Slides;
            dynamic slidesApi = slides;
            for (var slideIndex = 1; slideIndex <= Convert.ToInt32(slidesApi.Count); slideIndex++)
            {
                object? slide = null;
                object? shapes = null;
                try
                {
                    slide = slidesApi.Item(slideIndex);
                    shapes = ((dynamic)slide).Shapes;
                    dynamic shapesApi = shapes;
                    for (var shapeIndex = 1; shapeIndex <= Convert.ToInt32(shapesApi.Count); shapeIndex++)
                    {
                        object? shape = null;
                        try
                        {
                            shape = shapesApi.Item(shapeIndex);
                            if (!ReadTag(shape, "WENGGE_LINK_ID").Equals(linkId, StringComparison.Ordinal)) continue;
                            var result = (slide, shape);
                            slide = null;
                            shape = null;
                            return result;
                        }
                        finally { ComInterop.Release(shape); }
                    }
                }
                finally { ComInterop.Release(shapes); ComInterop.Release(slide); }
            }
            return (null, null);
        }
        finally { ComInterop.Release(slides); }
    }

    private static void SetPresentationTitle(dynamic presentation, dynamic slide, dynamic shapes, string linkId, string title)
    {
        object? titleShape = null;
        try
        {
            for (var index = 1; index <= Convert.ToInt32(shapes.Count); index++)
            {
                object? candidate = null;
                try
                {
                    candidate = shapes.Item(index);
                    if (!ReadTag(candidate, "WENGGE_MANAGED_TITLE").Equals(linkId, StringComparison.Ordinal)) continue;
                    titleShape = candidate;
                    candidate = null;
                    break;
                }
                finally { ComInterop.Release(candidate); }
            }
            titleShape ??= shapes.AddTextbox(1, 30, 20, Convert.ToDouble(presentation.PageSetup.SlideWidth) - 60, 40);
            SetTag(titleShape, "WENGGE_MANAGED_TITLE", linkId);
            dynamic textRange = ((dynamic)titleShape).TextFrame.TextRange;
            textRange.Text = title;
            textRange.Font.Size = 24;
            textRange.Font.Bold = true;
        }
        finally { ComInterop.Release(titleShape); }
    }

    private static string ReadTag(object owner, string name)
    {
        object? tags = null;
        try { tags = ((dynamic)owner).Tags; return Convert.ToString(((dynamic)tags).Item(name)) ?? string.Empty; }
        catch { return string.Empty; }
        finally { ComInterop.Release(tags); }
    }

    private static void SetTag(object owner, string name, string value)
    {
        object? tags = null;
        try
        {
            tags = ((dynamic)owner).Tags;
            try { ((dynamic)tags).Delete(name); } catch { }
            ((dynamic)tags).Add(name, value);
        }
        finally { ComInterop.Release(tags); }
    }

    private static void PrepareOutput(string output, bool updateExisting, bool overwrite)
    {
        if (!File.Exists(output)) return;
        if (updateExisting) return;
        if (!overwrite) throw new OfficeWorkerException("file_exists", $"输出文件已存在，请设置 params.overwrite=true: {output}");
        File.Delete(output);
    }

    private static string OutputPath(OfficeActionRequest request, string path, string extension)
    {
        var output = Path.GetFullPath(path);
        if (Path.GetExtension(output).Length == 0) output += extension;
        Directory.CreateDirectory(Path.GetDirectoryName(output) ?? Environment.CurrentDirectory);
        return output;
    }

    private static object Done(OfficeActionRequest request, string kind, string summary, object data, string? outputPath) =>
        OfficeActionResults.Done(request, "com", summary, data, [new OfficeChange(kind, outputPath, summary)], outputPath);

    private static string TargetInstanceId(OfficeActionRequest request, string app)
    {
        var specific = request.StringParam(app == "word" ? "wordInstanceId" : "presentationInstanceId");
        if (specific.Length > 0) return specific;
        return request.Operation == "buildReportPackage" ? string.Empty : request.StringParam("instanceId");
    }

    private sealed class TargetApplication : IDisposable
    {
        private readonly OfficeApplicationHandle? handle;
        private readonly OfficeDocumentLease? lease;

        private TargetApplication(OfficeApplicationHandle handle)
        {
            this.handle = handle;
            Application = handle.Application;
            ProgId = handle.ProgId;
            if (handle.ProgId.Contains("PowerPoint", StringComparison.OrdinalIgnoreCase)
                || handle.ProgId.Contains("wpp", StringComparison.OrdinalIgnoreCase))
            {
                try { ((dynamic)Application).Visible = -1; ((dynamic)Application).DisplayAlerts = 1; } catch { }
            }
            else { try { ((dynamic)Application).Visible = false; ((dynamic)Application).DisplayAlerts = 0; } catch { } }
        }

        private TargetApplication(OfficeDocumentLease lease)
        {
            this.lease = lease;
            Application = lease.Handle.Application;
            Document = lease.Handle.Document;
            ProgId = lease.Handle.ProgId;
        }

        public object Application { get; }
        public object? Document { get; }
        public string ProgId { get; }

        public static TargetApplication Open(OfficeApplicationProvider applications, string app, string host, string instanceId, string filePath)
        {
            if (instanceId.Length > 0)
            {
                var lease = OfficeDocumentService.AcquireDocument(app, filePath, instanceId);
                try { OfficeHostRouting.Validate(app, host, lease.Handle.ProgId); return new TargetApplication(lease); }
                catch { lease.Dispose(); throw; }
            }
            return new TargetApplication(applications.Create(OfficeHostRouting.ProgIds(app, host), $"未找到可用的 {app} Office/WPS COM 应用"));
        }

        public void Dispose()
        {
            lease?.Dispose();
            if (handle?.Created == true) { try { ((dynamic)Application).Quit(); } catch { } }
            handle?.Dispose();
            _ = handle?.WaitForExit();
        }
    }

    private sealed record WordBookmarkSnapshot(string Name, int Start, int End, string Text);

    internal sealed record ReportSection(string LinkId, JsonElement Values);

    private sealed record ShapeGeometry(double Left, double Top, double Width, double Height, double Rotation, int LockAspectRatio, int ZOrder)
    {
        public static ShapeGeometry Read(object shape)
        {
            dynamic api = shape;
            return new ShapeGeometry(
                Convert.ToDouble(api.Left), Convert.ToDouble(api.Top), Convert.ToDouble(api.Width), Convert.ToDouble(api.Height),
                Safe(() => Convert.ToDouble(api.Rotation), 0d), Safe(() => Convert.ToInt32(api.LockAspectRatio), 0), Safe(() => Convert.ToInt32(api.ZOrderPosition), 1));
        }

        public void Apply(dynamic shape)
        {
            try { shape.LockAspectRatio = 0; } catch { }
            shape.Left = Left;
            shape.Top = Top;
            shape.Width = Width;
            shape.Height = Height;
            try { shape.Rotation = Rotation; } catch { }
            try { shape.LockAspectRatio = LockAspectRatio; } catch { }
            for (var attempt = 0; attempt < 100; attempt++)
            {
                var current = Safe(() => Convert.ToInt32(shape.ZOrderPosition), ZOrder);
                if (current == ZOrder) break;
                shape.ZOrder(current > ZOrder ? 3 : 2);
            }
        }

        private static T Safe<T>(Func<T> value, T fallback) { try { return value(); } catch { return fallback; } }
    }
}
