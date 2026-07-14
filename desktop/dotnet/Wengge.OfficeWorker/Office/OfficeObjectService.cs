using System.Text.Json;
using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Protocol;
using Wengge.OfficeWorker.Runtime;

namespace Wengge.OfficeWorker.Office;

internal sealed class OfficeObjectService(OfficeDocumentService documents)
{
    public object ListObjects(JsonElement parameters)
    {
        var app = parameters.RequiredString("app");
        var filePath = parameters.RequiredString("filePath");
        var instanceId = parameters.OptionalString("instanceId");
        var kindFilter = parameters.OptionalString("kind");
        return documents.WithDocument(
            app, filePath, instanceId, null, 0,
            handle => app switch
            {
                "excel" => ListExcel(handle, kindFilter),
                "word" => ListWord(handle, kindFilter),
                "presentation" => ListPresentation(handle, kindFilter),
                _ => throw new OfficeWorkerException("unsupported_app", $"不支持的 Office 应用: {app}"),
            });
    }

    public object ActivateObject(JsonElement parameters)
    {
        var app = parameters.RequiredString("app");
        var filePath = parameters.RequiredString("filePath");
        var instanceId = parameters.OptionalString("instanceId");
        var locator = parameters.RequiredString("locator");
        return documents.WithDocument(
            app, filePath, instanceId, null, 0,
            handle =>
            {
                var selected = app switch
                {
                    "excel" => ActivateExcel(handle, locator),
                    "word" => ActivateWord(handle, locator),
                    "presentation" => ActivatePresentation(handle, locator),
                    _ => throw new OfficeWorkerException("unsupported_app", $"不支持的 Office 应用: {app}"),
                };
                if (app != "presentation") ((dynamic)handle.Document).Activate();
                return Entry(handle, selected.Kind, selected.Name, locator, selected.Parent, selected.Index, string.Empty, true);
            });
    }

    private static object[] ListExcel(OfficeDocumentHandle handle, string? filter)
    {
        dynamic workbook = handle.Document;
        dynamic app = handle.Application;
        var result = new List<object>();
        var activeSheetName = OfficeDocumentService.SafeString(() => workbook.ActiveSheet.Name);
        ForEach(() => workbook.Worksheets, (sheet, sheetIndex) =>
        {
            var sheetName = OfficeDocumentService.SafeString(() => sheet.Name);
            var sheetPart = OfficeDocumentService.EncodeLocator(sheetName);
            Add(result, handle, filter, "sheet", sheetName, $"sheet:{sheetPart}", index: OfficeDocumentService.SafeInt32(() => sheet.Index, sheetIndex), selected: sheetName == activeSheetName);
            object? used = null;
            try
            {
                used = sheet.UsedRange;
                var address = OfficeDocumentService.SafeString(() => ((dynamic)used).Address(false, false));
                Add(result, handle, filter, "range", address, $"range:{sheetPart}/{OfficeDocumentService.EncodeLocator(address)}", sheetName, 1,
                    $"rows={OfficeDocumentService.SafeInt32(() => ((dynamic)used).Rows.Count)};columns={OfficeDocumentService.SafeInt32(() => ((dynamic)used).Columns.Count)}");
            }
            catch { }
            finally { ComInterop.Release(used); }

            ForEach(() => sheet.ListObjects, (table, index) =>
            {
                var name = OfficeDocumentService.SafeString(() => table.Name);
                Add(result, handle, filter, "table", name, $"table:{sheetPart}/{OfficeDocumentService.EncodeLocator(name)}", sheetName,
                    OfficeDocumentService.SafeInt32(() => table.Index, index), OfficeDocumentService.SafeString(() => table.Range.Address(false, false)));
            });
            ForEach(() => sheet.ChartObjects(), (chart, index) =>
            {
                var name = OfficeDocumentService.SafeString(() => chart.Name);
                Add(result, handle, filter, "chart", name, $"chart:{sheetPart}/{OfficeDocumentService.EncodeLocator(name)}", sheetName,
                    OfficeDocumentService.SafeInt32(() => chart.Index, index));
            });
            ForEach(() => sheet.Shapes, (shape, index) =>
            {
                var name = OfficeDocumentService.SafeString(() => shape.Name);
                Add(result, handle, filter, "shape", name, $"shape:{sheetPart}/{OfficeDocumentService.EncodeLocator(name)}", sheetName,
                    OfficeDocumentService.SafeInt32(() => shape.Id, index));
            });
            ForEach(() => sheet.PivotTables(), (pivot, _) =>
            {
                var name = OfficeDocumentService.SafeString(() => pivot.Name);
                Add(result, handle, filter, "pivotTable", name, $"pivotTable:{sheetPart}/{OfficeDocumentService.EncodeLocator(name)}", sheetName, 0,
                    OfficeDocumentService.SafeString(() => pivot.TableRange2.Address(false, false)));
            });
        });

        TryAddExcelSelection(result, handle, filter, app, workbook, "ActiveCell", "cell");
        TryAddExcelSelection(result, handle, filter, app, workbook, "Selection", "range");
        ForEach(() => workbook.Names, (name, index) =>
        {
            var value = OfficeDocumentService.SafeString(() => name.Name);
            Add(result, handle, filter, "name", value, $"name:{OfficeDocumentService.EncodeLocator(value)}", index: OfficeDocumentService.SafeInt32(() => name.Index, index),
                detail: OfficeDocumentService.SafeString(() => name.RefersTo));
        });
        ForEach(() => workbook.Queries, (query, _) =>
        {
            var name = OfficeDocumentService.SafeString(() => query.Name);
            Add(result, handle, filter, "query", name, $"query:{OfficeDocumentService.EncodeLocator(name)}", detail: OfficeDocumentService.SafeString(() => query.Formula));
        });
        ForEach(() => workbook.Connections, (connection, _) =>
        {
            var name = OfficeDocumentService.SafeString(() => connection.Name);
            Add(result, handle, filter, "connection", name, $"connection:{OfficeDocumentService.EncodeLocator(name)}", detail: OfficeDocumentService.SafeString(() => connection.Description));
        });
        ForEach(() => workbook.SlicerCaches, (cache, _) =>
        {
            var cacheName = OfficeDocumentService.SafeString(() => cache.Name);
            ForEach(() => cache.Slicers, (slicer, _) =>
            {
                var name = OfficeDocumentService.SafeString(() => slicer.Name);
                Add(result, handle, filter, "slicer", name,
                    $"slicer:{OfficeDocumentService.EncodeLocator(cacheName)}/{OfficeDocumentService.EncodeLocator(name)}", cacheName,
                    detail: OfficeDocumentService.SafeString(() => slicer.Caption));
            });
        });
        return result.ToArray();
    }

    private static object[] ListWord(OfficeDocumentHandle handle, string? filter)
    {
        dynamic document = handle.Document;
        var result = new List<object>();
        var pageCount = OfficeDocumentService.SafeInt32(() => document.ComputeStatistics(2));
        for (var page = 1; page <= pageCount; page++) Add(result, handle, filter, "page", $"第 {page} 页", $"page:{page}", index: page);
        ForEach(() => document.Sections, (_, index) => Add(result, handle, filter, "section", $"第 {index} 节", $"section:{index}", index: index));
        ForEach(() => document.Paragraphs, (paragraph, _) =>
        {
            object? range = null;
            try
            {
                range = paragraph.Range;
                dynamic rangeApi = range;
                var start = OfficeDocumentService.SafeInt32(() => rangeApi.Start);
                var text = Truncate(OfficeDocumentService.SafeString(() => rangeApi.Text).Trim());
                Add(result, handle, filter, "paragraph", text.Length > 0 ? text : $"空段落 {start}", $"paragraph:{start}", index: start, detail: text);
                var outlineLevel = OfficeDocumentService.SafeInt32(() => paragraph.OutlineLevel, 10);
                if (outlineLevel is >= 1 and <= 9)
                    Add(result, handle, filter, "heading", text.Length > 0 ? text : $"标题 {start}", $"heading:{start}", index: start, detail: $"level={outlineLevel}");
                var style = OfficeDocumentService.SafeString(() => rangeApi.Style.NameLocal, OfficeDocumentService.SafeString(() => rangeApi.Style));
                if (style.Contains("caption", StringComparison.OrdinalIgnoreCase) || style.Contains("题注", StringComparison.OrdinalIgnoreCase))
                    Add(result, handle, filter, "caption", text.Length > 0 ? text : $"题注 {start}", $"caption:{start}", index: start, detail: style);
            }
            finally { ComInterop.Release(range); }
        });
        ForEach(() => document.Tables, (table, index) => Add(result, handle, filter, "table", $"表格 {index}", $"table:{index}", index: index,
            detail: $"rows={OfficeDocumentService.SafeInt32(() => table.Rows.Count)};columns={OfficeDocumentService.SafeInt32(() => table.Columns.Count)}"));
        ForEach(() => document.Bookmarks, (bookmark, _) =>
        {
            var name = OfficeDocumentService.SafeString(() => bookmark.Name);
            Add(result, handle, filter, "bookmark", name, $"bookmark:{OfficeDocumentService.EncodeLocator(name)}", index: OfficeDocumentService.SafeInt32(() => bookmark.Range.Start));
        });
        ForEach(() => document.ContentControls, (control, _) =>
        {
            var id = OfficeDocumentService.SafeInt64(() => control.ID);
            var title = OfficeDocumentService.SafeString(() => control.Title);
            if (title.Length == 0) title = OfficeDocumentService.SafeString(() => control.Tag, id.ToString(System.Globalization.CultureInfo.InvariantCulture));
            Add(result, handle, filter, "contentControl", title, $"contentControl:{id}", index: ClampIndex(id), detail: OfficeDocumentService.SafeString(() => control.Tag));
        });
        AddIndexed(result, handle, filter, () => document.InlineShapes, "inlineShape", "嵌入对象", "inlineShape");
        ForEach(() => document.Shapes, (shape, index) =>
        {
            var name = OfficeDocumentService.SafeString(() => shape.Name);
            Add(result, handle, filter, "shape", name, $"shape:{OfficeDocumentService.EncodeLocator(name)}", index: OfficeDocumentService.SafeInt32(() => shape.ID, index));
        });
        ForEach(() => document.Comments, (comment, index) => Add(result, handle, filter, "comment",
            Truncate(OfficeDocumentService.SafeString(() => comment.Range.Text).Trim(), $"批注 {index}"), $"comment:{index}", index: index,
            detail: OfficeDocumentService.SafeString(() => comment.Author)));
        ForEach(() => document.Revisions, (revision, index) => Add(result, handle, filter, "revision", $"修订 {index}", $"revision:{index}", index: index,
            detail: $"type={OfficeDocumentService.SafeInt32(() => revision.Type)};author={OfficeDocumentService.SafeString(() => revision.Author)}"));
        ForEach(() => document.Footnotes, (note, index) => Add(result, handle, filter, "footnote", $"脚注 {index}", $"footnote:{index}", index: index,
            detail: OfficeDocumentService.SafeString(() => note.Range.Text).Trim()));
        ForEach(() => document.Endnotes, (note, index) => Add(result, handle, filter, "endnote", $"尾注 {index}", $"endnote:{index}", index: index,
            detail: OfficeDocumentService.SafeString(() => note.Range.Text).Trim()));
        return result.ToArray();
    }

    private static object[] ListPresentation(OfficeDocumentHandle handle, string? filter)
    {
        dynamic presentation = handle.Document;
        dynamic app = handle.Application;
        var result = new List<object>();
        var activeSlideIndex = OfficeDocumentService.SafeInt32(() => app.ActiveWindow.View.Slide.SlideIndex);
        ForEach(() => presentation.Designs, (design, designIndex) =>
        {
            var name = OfficeDocumentService.SafeString(() => design.Name);
            Add(result, handle, filter, "master", name, $"master:{designIndex}", index: designIndex,
                detail: OfficeDocumentService.SafeString(() => design.SlideMaster.Name));
            ForEach(() => design.SlideMaster.CustomLayouts, (layout, layoutIndex) =>
            {
                var layoutName = OfficeDocumentService.SafeString(() => layout.Name);
                Add(result, handle, filter, "layout", layoutName, $"layout:{designIndex}/{layoutIndex}", $"master:{designIndex}", layoutIndex);
            });
        });
        ForEach(() => presentation.Slides, (slide, _) =>
        {
            var slideIndex = OfficeDocumentService.SafeInt32(() => slide.SlideIndex);
            var slideId = OfficeDocumentService.SafeInt32(() => slide.SlideID);
            var title = OfficeDocumentService.SafeString(() => slide.Shapes.Title.TextFrame.TextRange.Text);
            Add(result, handle, filter, "slide", title.Length > 0 ? title : $"幻灯片 {slideIndex}", $"slide:{slideId}", index: slideIndex,
                detail: $"slideId={slideId};title={title}", selected: slideIndex == activeSlideIndex);
            Add(result, handle, filter, "notesPage", $"幻灯片 {slideIndex} 备注页", $"notesPage:{slideId}", $"slide:{slideId}", slideIndex);
            ForEach(() => slide.Shapes, (shape, _) => AddPresentationShape(result, handle, filter, shape, slideId, $"slide:{slideId}", string.Empty));
        });
        return result.ToArray();
    }

    private static SelectedObject ActivateExcel(OfficeDocumentHandle handle, string locator)
    {
        dynamic workbook = handle.Document;
        if (TryLocator(locator, "sheet", out var sheetValue))
        {
            var sheetName = OfficeDocumentService.DecodeLocator(sheetValue);
            dynamic sheet = workbook.Worksheets.Item(sheetName); sheet.Activate();
            return new("sheet", sheetName, string.Empty, OfficeDocumentService.SafeInt32(() => sheet.Index));
        }
        foreach (var kind in new[] { "cell", "range", "table", "chart", "shape", "pivotTable" })
        {
            if (!TryLocator(locator, kind, out var value)) continue;
            var parts = value.Split('/', 2);
            if (parts.Length != 2) throw new OfficeWorkerException("invalid_locator", "Excel 对象 locator 格式无效");
            var sheetName = OfficeDocumentService.DecodeLocator(parts[0]);
            var objectName = OfficeDocumentService.DecodeLocator(parts[1]);
            dynamic sheet = workbook.Worksheets.Item(sheetName); sheet.Activate();
            return kind switch
            {
                "cell" or "range" => SelectExcelRange(sheet, kind, objectName, sheetName),
                "table" => SelectExcelTable(sheet, objectName, sheetName),
                "chart" => SelectExcelChart(sheet, objectName, sheetName),
                "shape" => SelectExcelShape(sheet, objectName, sheetName),
                _ => SelectExcelPivot(sheet, objectName, sheetName),
            };
        }
        if (TryLocator(locator, "name", out var nameValue))
        {
            var name = OfficeDocumentService.DecodeLocator(nameValue); dynamic item = workbook.Names.Item(name); item.RefersToRange.Select();
            return new("name", name, string.Empty, OfficeDocumentService.SafeInt32(() => item.Index));
        }
        if (TryLocator(locator, "slicer", out var slicerValue))
        {
            var parts = slicerValue.Split('/', 2); if (parts.Length != 2) throw new OfficeWorkerException("invalid_locator", "切片器 locator 格式无效");
            var cacheName = OfficeDocumentService.DecodeLocator(parts[0]); var name = OfficeDocumentService.DecodeLocator(parts[1]);
            dynamic slicer = workbook.SlicerCaches.Item(cacheName).Slicers.Item(name); slicer.Shape.Select();
            return new("slicer", name, cacheName, 0);
        }
        foreach (var kind in new[] { "query", "connection" })
        {
            if (!TryLocator(locator, kind, out var value)) continue;
            var name = OfficeDocumentService.DecodeLocator(value); workbook.Activate();
            LocateQueryTable(workbook, kind, name);
            return new(kind, name, string.Empty, 0);
        }
        throw UnsupportedLocator(locator);
    }

    private static SelectedObject ActivateWord(OfficeDocumentHandle handle, string locator)
    {
        dynamic document = handle.Document; dynamic app = handle.Application;
        if (TryPositiveIndex(locator, "page", out var index)) { app.Selection.GoTo(1, 1, index); return new("page", $"第 {index} 页", string.Empty, index); }
        if (TryPositiveIndex(locator, "section", out index)) { document.Sections.Item(index).Range.Select(); return new("section", $"第 {index} 节", string.Empty, index); }
        foreach (var kind in new[] { "paragraph", "heading", "caption" })
        {
            if (!TryPositiveIndex(locator, kind, out index)) continue;
            string name = string.Empty;
            ForEach(() => document.Paragraphs, (paragraph, _) =>
            {
                if (OfficeDocumentService.SafeInt32(() => paragraph.Range.Start) != index) return;
                paragraph.Range.Select(); name = OfficeDocumentService.SafeString(() => paragraph.Range.Text).Trim();
            });
            if (name.Length == 0) throw UnsupportedLocator(locator);
            return new(kind, name, string.Empty, index);
        }
        if (TryPositiveIndex(locator, "table", out index)) { document.Tables.Item(index).Range.Select(); return new("table", $"表格 {index}", string.Empty, index); }
        if (TryLocator(locator, "bookmark", out var value)) { var name = OfficeDocumentService.DecodeLocator(value); document.Bookmarks.Item(name).Range.Select(); return new("bookmark", name, string.Empty, 0); }
        if (TryPositiveIndex(locator, "contentControl", out index))
        {
            string name = string.Empty;
            ForEach(() => document.ContentControls, (control, _) =>
            {
                if (OfficeDocumentService.SafeInt64(() => control.ID) != index) return;
                control.Range.Select(); name = OfficeDocumentService.SafeString(() => control.Title, index.ToString(System.Globalization.CultureInfo.InvariantCulture));
            });
            if (name.Length == 0) throw UnsupportedLocator(locator);
            return new("contentControl", name, string.Empty, index);
        }
        if (TryPositiveIndex(locator, "inlineShape", out index)) { document.InlineShapes.Item(index).Range.Select(); return new("inlineShape", $"嵌入对象 {index}", string.Empty, index); }
        if (TryLocator(locator, "shape", out value)) { var name = OfficeDocumentService.DecodeLocator(value); dynamic shape = document.Shapes.Item(name); shape.Select(); return new("shape", name, string.Empty, OfficeDocumentService.SafeInt32(() => shape.ID)); }
        foreach (var kind in new[] { "comment", "revision", "footnote", "endnote" })
        {
            if (!TryPositiveIndex(locator, kind, out index)) continue;
            dynamic item = kind switch
            {
                "comment" => document.Comments.Item(index),
                "revision" => document.Revisions.Item(index),
                "footnote" => document.Footnotes.Item(index),
                _ => document.Endnotes.Item(index),
            };
            if (kind == "comment") item.Scope.Select(); else if (kind is "footnote" or "endnote") item.Reference.Select(); else item.Range.Select();
            var label = kind switch { "comment" => OfficeDocumentService.SafeString(() => item.Range.Text).Trim(), "revision" => $"修订 {index}", "footnote" => $"脚注 {index}", _ => $"尾注 {index}" };
            return new(kind, label, string.Empty, index);
        }
        throw UnsupportedLocator(locator);
    }

    private static SelectedObject ActivatePresentation(OfficeDocumentHandle handle, string locator)
    {
        dynamic presentation = handle.Document; dynamic app = handle.Application;
        if (TryPositiveIndex(locator, "slide", out var slideId))
        {
            dynamic slide = SlideById(presentation, slideId); ActivatePresentationSlide(presentation, app, slide); var index = OfficeDocumentService.SafeInt32(() => slide.SlideIndex);
            return new("slide", $"幻灯片 {index}", string.Empty, index);
        }
        if (TryPositiveIndex(locator, "notesPage", out slideId))
        {
            dynamic slide = SlideById(presentation, slideId); var index = OfficeDocumentService.SafeInt32(() => slide.SlideIndex);
            app.ActiveWindow.View.GotoSlide(index); app.ActiveWindow.ViewType = 3;
            return new("notesPage", $"幻灯片 {index} 备注页", $"slide:{slideId}", index);
        }
        if (TryPositiveIndex(locator, "master", out var designIndex))
        {
            dynamic master = presentation.Designs.Item(designIndex).SlideMaster; app.ActiveWindow.ViewType = 2; try { master.Select(); } catch { }
            return new("master", OfficeDocumentService.SafeString(() => master.Name), string.Empty, designIndex);
        }
        if (TryLocator(locator, "layout", out var layoutValue))
        {
            var parts = layoutValue.Split('/', 2); if (parts.Length != 2 || !int.TryParse(parts[0], out designIndex) || !int.TryParse(parts[1], out var layoutIndex)) throw new OfficeWorkerException("invalid_locator", "版式 locator 格式无效");
            dynamic layout = presentation.Designs.Item(designIndex).SlideMaster.CustomLayouts.Item(layoutIndex); app.ActiveWindow.ViewType = 2;
            try { layout.Select(); } catch { try { presentation.Designs.Item(designIndex).SlideMaster.Select(); } catch { } }
            return new("layout", OfficeDocumentService.SafeString(() => layout.Name), $"master:{designIndex}", layoutIndex);
        }
        foreach (var kind in new[] { "shape", "chart", "table" })
        {
            if (!TryLocator(locator, kind, out var value)) continue;
            var separator = value.IndexOf('/');
            if (separator < 1 || !int.TryParse(value[..separator], out slideId)) throw new OfficeWorkerException("invalid_locator", "形状 locator 格式无效");
            dynamic slide = SlideById(presentation, slideId); ActivatePresentationSlide(presentation, app, slide);
            dynamic shape = ShapeByPath(slide, value[(separator + 1)..].Split('/'));
            shape.Select();
            return new(kind, OfficeDocumentService.SafeString(() => shape.Name), $"slide:{slideId}", OfficeDocumentService.SafeInt32(() => shape.Id));
        }
        throw UnsupportedLocator(locator);
    }

    private static void ActivatePresentationSlide(dynamic presentation, dynamic app, dynamic slide)
    {
        object? windows = null;
        object? window = null;
        object? view = null;
        try
        {
            try { app.Visible = -1; } catch { }
            windows = presentation.Windows;
            window = ((dynamic)windows).Item(1);
            ((dynamic)window).Activate();
            try { ((dynamic)window).ViewType = 9; } catch { }
            view = ((dynamic)window).View;
            ((dynamic)view).GotoSlide(Convert.ToInt32(slide.SlideIndex));
            slide.Select();
        }
        finally
        {
            ComInterop.Release(view);
            ComInterop.Release(window);
            ComInterop.Release(windows);
        }
    }

    private static void TryAddExcelSelection(List<object> result, OfficeDocumentHandle handle, string? filter, dynamic app, dynamic workbook, string property, string kind)
    {
        object? selection = null;
        try
        {
            selection = app.GetType().InvokeMember(property, System.Reflection.BindingFlags.GetProperty, null, app, null);
            if (selection is null) return;
            dynamic item = selection;
            if (!string.Equals(OfficeDocumentService.PathKey(OfficeDocumentService.SafeString(() => item.Parent.Parent.FullName)), OfficeDocumentService.PathKey(OfficeDocumentService.SafeString(() => workbook.FullName)), StringComparison.OrdinalIgnoreCase)) return;
            var sheetName = OfficeDocumentService.SafeString(() => item.Worksheet.Name);
            var address = OfficeDocumentService.SafeString(() => item.Address(false, false));
            Add(result, handle, filter, kind, address, $"{kind}:{OfficeDocumentService.EncodeLocator(sheetName)}/{OfficeDocumentService.EncodeLocator(address)}", sheetName,
                detail: property == "Selection" ? "当前选区" : string.Empty, selected: true);
        }
        catch { }
        finally { ComInterop.Release(selection); }
    }

    private static void AddPresentationShape(List<object> result, OfficeDocumentHandle handle, string? filter, dynamic shape, int slideId, string parent, string path)
    {
        var name = OfficeDocumentService.SafeString(() => shape.Name);
        var part = OfficeDocumentService.EncodeLocator(name);
        var shapePath = path.Length > 0 ? $"{path}/{part}" : part;
        var detail = Truncate(OfficeDocumentService.SafeString(() => shape.TextFrame.TextRange.Text));
        Add(result, handle, filter, "shape", name, $"shape:{slideId}/{shapePath}", parent, OfficeDocumentService.SafeInt32(() => shape.Id), detail);
        if (OfficeDocumentService.SafeBoolean(() => shape.HasChart, false)) Add(result, handle, filter, "chart", name, $"chart:{slideId}/{shapePath}", parent, OfficeDocumentService.SafeInt32(() => shape.Id));
        if (OfficeDocumentService.SafeBoolean(() => shape.HasTable, false)) Add(result, handle, filter, "table", name, $"table:{slideId}/{shapePath}", parent, OfficeDocumentService.SafeInt32(() => shape.Id));
        if (OfficeDocumentService.SafeInt32(() => shape.Type) == 6)
            ForEach(() => shape.GroupItems, (child, _) => AddPresentationShape(result, handle, filter, child, slideId, $"shape:{slideId}/{shapePath}", shapePath));
    }

    private static void AddIndexed(List<object> result, OfficeDocumentHandle handle, string? filter, Func<object?> collection, string kind, string label, string locator)
    {
        ForEach(collection, (_, index) => Add(result, handle, filter, kind, $"{label} {index}", $"{locator}:{index}", index: index));
    }

    private static void Add(List<object> result, OfficeDocumentHandle handle, string? filter, string kind, string name, string locator,
        string parent = "", int index = 0, string detail = "", bool selected = false)
    {
        if (!string.IsNullOrWhiteSpace(filter) && filter != kind) return;
        result.Add(Entry(handle, kind, name, locator, parent, index, detail, selected));
    }

    private static object Entry(OfficeDocumentHandle handle, string kind, string name, string locator, string parent, int index, string detail, bool selected) => new
    {
        app = handle.App,
        documentPath = OfficeDocumentService.SafeString(() => ((dynamic)handle.Document).FullName),
        instanceId = handle.InstanceId,
        kind,
        name,
        locator,
        parent,
        index,
        detail,
        selected,
    };

    private static void ForEach(Func<object?> collectionFactory, Action<dynamic, int> action)
    {
        object? collection = null;
        try
        {
            collection = collectionFactory();
            if (collection is null) return;
            dynamic items = collection;
            var count = OfficeDocumentService.SafeInt32(() => items.Count);
            for (var index = 1; index <= count; index++)
            {
                object? item = null;
                try { item = items.Item(index); if (item is not null) action((dynamic)item, index); }
                catch { }
                finally { ComInterop.Release(item); }
            }
        }
        catch { }
        finally { ComInterop.Release(collection); }
    }

    private static SelectedObject SelectExcelRange(dynamic sheet, string kind, string address, string parent)
    {
        dynamic range = sheet.Range(address); range.Select(); return new(kind, address, parent, 0);
    }
    private static SelectedObject SelectExcelTable(dynamic sheet, string name, string parent)
    {
        dynamic item = sheet.ListObjects.Item(name); item.Range.Select(); return new("table", OfficeDocumentService.SafeString(() => item.Name), parent, OfficeDocumentService.SafeInt32(() => item.Index));
    }
    private static SelectedObject SelectExcelChart(dynamic sheet, string name, string parent)
    {
        dynamic item = sheet.ChartObjects(name); item.Activate(); return new("chart", OfficeDocumentService.SafeString(() => item.Name), parent, OfficeDocumentService.SafeInt32(() => item.Index));
    }
    private static SelectedObject SelectExcelShape(dynamic sheet, string name, string parent)
    {
        dynamic item = sheet.Shapes.Item(name); item.Select(); return new("shape", OfficeDocumentService.SafeString(() => item.Name), parent, OfficeDocumentService.SafeInt32(() => item.Id));
    }
    private static SelectedObject SelectExcelPivot(dynamic sheet, string name, string parent)
    {
        dynamic item = sheet.PivotTables(name); item.TableRange2.Select(); return new("pivotTable", OfficeDocumentService.SafeString(() => item.Name), parent, 0);
    }

    private static void LocateQueryTable(dynamic workbook, string kind, string name)
    {
        ForEach(() => workbook.Worksheets, (sheet, _) => ForEach(() => sheet.ListObjects, (table, _) =>
        {
            var connection = OfficeDocumentService.SafeString(() => table.QueryTable.WorkbookConnection.Name);
            var matches = kind == "connection" ? string.Equals(connection, name, StringComparison.OrdinalIgnoreCase)
                : string.Equals(connection, name, StringComparison.OrdinalIgnoreCase) || string.Equals(connection, $"Query - {name}", StringComparison.OrdinalIgnoreCase);
            if (matches) { sheet.Activate(); table.Range.Select(); }
        }));
    }

    private static dynamic SlideById(dynamic presentation, int slideId)
    {
        object? slides = null;
        try
        {
            slides = presentation.Slides;
            dynamic collection = slides;
            var count = OfficeDocumentService.SafeInt32(() => collection.Count);
            for (var index = 1; index <= count; index++)
            {
                object? slide = collection.Item(index);
                if (slide is not null && OfficeDocumentService.SafeInt32(() => ((dynamic)slide).SlideID) == slideId) return slide;
                ComInterop.Release(slide);
            }
        }
        finally { ComInterop.Release(slides); }
        throw new OfficeWorkerException("object_not_found", "找不到指定 SlideID");
    }

    private static dynamic ShapeByPath(dynamic slide, IEnumerable<string> encodedParts)
    {
        dynamic? shape = null;
        foreach (var encoded in encodedParts)
        {
            var name = OfficeDocumentService.DecodeLocator(encoded);
            shape = shape is null ? slide.Shapes.Item(name) : shape.GroupItems.Item(name);
        }
        return shape ?? throw new OfficeWorkerException("object_not_found", "找不到指定形状路径");
    }

    private static bool TryLocator(string locator, string kind, out string value)
    {
        var prefix = $"{kind}:";
        if (locator.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) { value = locator[prefix.Length..]; return value.Length > 0; }
        value = string.Empty; return false;
    }

    private static bool TryPositiveIndex(string locator, string kind, out int index)
    {
        index = 0;
        return TryLocator(locator, kind, out var value) && int.TryParse(value, out index) && index > 0;
    }

    private static OfficeWorkerException UnsupportedLocator(string locator) =>
        new("unsupported_locator", $"不支持或找不到 Office 对象 locator: {locator}");

    private static string Truncate(string value, string fallback = "")
    {
        if (value.Length == 0) return fallback;
        return value.Length > 120 ? value[..120] : value;
    }

    private static int ClampIndex(long value) => value > int.MaxValue ? int.MaxValue : value < int.MinValue ? int.MinValue : (int)value;
    private sealed record SelectedObject(string Kind, string Name, string Parent, int Index);
}
