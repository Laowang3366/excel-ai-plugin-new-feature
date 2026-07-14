using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Office;

internal sealed class WordMailMergeContentActionService(OfficeApplicationProvider applications)
{
    private static readonly HashSet<string> Operations =
    ["prepareMailMergeTemplate", "mailMerge", "batchMailMerge", "inspectContentControls", "populateContentControls", "manageContentControls"];

    public static bool Supports(string operation) => Operations.Contains(operation);

    public object Execute(OfficeActionRequest request)
    {
        using var context = new WordActionContext(applications, request);
        if (request.Operation == "inspectContentControls")
            return OfficeActionResults.Done(request, "com", "已检查内容控件", InspectControls(context), Array.Empty<OfficeChange>());
        if (request.Operation is "mailMerge" or "batchMailMerge") return Merge(context, request);
        var data = request.Operation == "prepareMailMergeTemplate" ? PrepareTemplate(context, request)
            : request.Operation == "populateContentControls" ? PopulateControls(context, request)
            : ManageControls(context, request);
        context.Save(request);
        return OfficeActionResults.Done(request, "com", "已更新 Word 模板字段", data,
            [new OfficeChange("word-template", request.Target, "已更新 Word 模板字段")]);
    }

    private static object InspectControls(WordActionContext context)
    {
        var controls = new List<object>();
        object? collection = null;
        try
        {
            collection = context.Document.ContentControls; dynamic collectionApi = collection;
            for (var index = 1; index <= Convert.ToInt32(collectionApi.Count); index++)
            {
                object? control = null; object? range = null;
                try
                {
                    control = collectionApi.Item(index); dynamic controlApi = control; range = controlApi.Range;
                    string rawTag = Convert.ToString((object)controlApi.Tag) ?? string.Empty;
                    var (tag, logicalType) = DecodeControlTag(rawTag);
                    var text = (Convert.ToString(((dynamic)range).Text) ?? string.Empty).TrimEnd('\r', '\a');
                    controls.Add(new
                    {
                        index,
                        id = Safe(() => controlApi.ID),
                        title = Safe(() => controlApi.Title),
                        tag,
                        type = Safe(() => controlApi.Type),
                        logicalType,
                        text,
                        @checked = logicalType == "checkbox" ? text == "[x]" : Safe(() => controlApi.Checked),
                        lockContents = Safe(() => controlApi.LockContents),
                        lockControl = Safe(() => controlApi.LockContentControl),
                    });
                }
                finally { ComInterop.Release(range); ComInterop.Release(control); }
            }
            return new { controls, controlCount = controls.Count };
        }
        finally { ComInterop.Release(collection); }
    }

    private static object PopulateControls(WordActionContext context, OfficeActionRequest request)
    {
        var values = request.Param("values");
        if (values.ValueKind != JsonValueKind.Object) throw new OfficeWorkerException("invalid_params", "populateContentControls 需要 params.values 对象");
        var populated = new List<string>();
        object? collection = null;
        try
        {
            collection = context.Document.ContentControls; dynamic collectionApi = collection;
            for (var index = 1; index <= Convert.ToInt32(collectionApi.Count); index++)
            {
                object? control = null; object? range = null;
                try
                {
                    control = collectionApi.Item(index); dynamic controlApi = control;
                    var title = Convert.ToString(controlApi.Title) ?? string.Empty;
                    string rawTag = Convert.ToString((object)controlApi.Tag) ?? string.Empty;
                    var (tag, logicalType) = DecodeControlTag(rawTag);
                    JsonElement value;
                    if (!values.TryGetProperty(tag, out value) && !values.TryGetProperty(title, out value)) continue;
                    range = controlApi.Range; dynamic rangeApi = range;
                    var type = Convert.ToInt32(controlApi.Type);
                    if (type == 8) controlApi.Checked = BooleanValue(value);
                    else if (logicalType == "checkbox") rangeApi.Text = BooleanValue(value) ? "[x]" : "[ ]";
                    else if (type == 6)
                    {
                        var dateFormat = value.ValueKind == JsonValueKind.Object && value.TryGetProperty("dateFormat", out var formatValue) ? formatValue.GetString() : request.StringParam("dateFormat", "yyyy-MM-dd");
                        controlApi.DateDisplayFormat = dateFormat ?? "yyyy-MM-dd";
                        rangeApi.Text = ElementValue(value);
                    }
                    else if (type is 3 or 4)
                    {
                        SelectListEntry(controlApi, ElementValue(value));
                    }
                    else if ((type == 2 || logicalType == "picture") && File.Exists(ElementValue(value)))
                    {
                        object? images = null; object? image = null;
                        try { images = rangeApi.InlineShapes; image = ((dynamic)images).AddPicture(ElementValue(value), false, true, range); }
                        finally { ComInterop.Release(image); ComInterop.Release(images); }
                    }
                    else rangeApi.Text = ElementValue(value);
                    populated.Add(tag.Length > 0 ? tag : title);
                }
                finally { ComInterop.Release(range); ComInterop.Release(control); }
            }
            return new { populated, populatedCount = populated.Count };
        }
        finally { ComInterop.Release(collection); }
    }

    private static object ManageControls(WordActionContext context, OfficeActionRequest request)
    {
        var command = request.StringParam("command", "add");
        object? collection = null; object? control = null; object? range = null;
        try
        {
            collection = context.Document.ContentControls; dynamic collectionApi = collection;
            if (command == "add")
            {
                var controls = request.Param("controls");
                if (controls.ValueKind == JsonValueKind.Array)
                {
                    var added = new List<object>();
                    foreach (var item in controls.EnumerateArray()) added.Add(AddControl(context, collectionApi, item));
                    return new { command, added, addedCount = added.Count };
                }
                range = ResolveRange(context, request);
                control = collectionApi.Add(ControlType(request.StringParam("type", "text")), range);
                ConfigureControl(control, request.Params);
            }
            else
            {
                control = FindControl(collectionApi, request);
                if (control is null) throw new OfficeWorkerException("content_control_not_found", "找不到指定内容控件");
                dynamic controlApi = control;
                if (command == "delete") controlApi.Delete(request.BoolParam("deleteContents"));
                else if (command == "update")
                {
                    if (request.Param("title").ValueKind == JsonValueKind.String) controlApi.Title = request.StringParam("title");
                    if (request.Param("tag").ValueKind == JsonValueKind.String) controlApi.Tag = request.StringParam("tag");
                    if (request.Param("lockContents").ValueKind is JsonValueKind.True or JsonValueKind.False) controlApi.LockContents = request.BoolParam("lockContents");
                    if (request.Param("lockControl").ValueKind is JsonValueKind.True or JsonValueKind.False) controlApi.LockContentControl = request.BoolParam("lockControl");
                }
                else throw new OfficeWorkerException("unsupported_operation", $"不支持的内容控件命令: {command}");
            }
            return new { command, title = request.StringParam("title"), tag = request.StringParam("tag") };
        }
        finally { ComInterop.Release(range); ComInterop.Release(control); ComInterop.Release(collection); }
    }

    private static object PrepareTemplate(WordActionContext context, OfficeActionRequest request)
    {
        var fields = request.Param("fields");
        if (fields.ValueKind != JsonValueKind.Array) throw new OfficeWorkerException("invalid_params", "prepareMailMergeTemplate 需要 params.fields");
        object? mailMerge = null; object? mergeFields = null;
        var added = new List<string>();
        try
        {
            mailMerge = context.Document.MailMerge; mergeFields = ((dynamic)mailMerge).Fields; dynamic fieldsApi = mergeFields;
            foreach (var field in fields.EnumerateArray())
            {
                var name = field.ValueKind == JsonValueKind.String ? field.GetString() : field.TryGetProperty("name", out var fieldName) ? fieldName.GetString() : null;
                if (string.IsNullOrWhiteSpace(name)) continue;
                object? range = null; object? mergeField = null;
                try
                {
                    range = context.Document.Range(Math.Max(0, Convert.ToInt32(context.Document.Content.End) - 1), Math.Max(0, Convert.ToInt32(context.Document.Content.End) - 1));
                    ((dynamic)range).InsertAfter(name + ": ");
                    range = context.Document.Range(Math.Max(0, Convert.ToInt32(context.Document.Content.End) - 1), Math.Max(0, Convert.ToInt32(context.Document.Content.End) - 1));
                    mergeField = fieldsApi.Add(range, name);
                    ((dynamic)range).InsertAfter(Environment.NewLine);
                    added.Add(name);
                }
                finally { ComInterop.Release(mergeField); ComInterop.Release(range); }
            }
            return new { fields = added, fieldCount = added.Count };
        }
        finally { ComInterop.Release(mergeFields); ComInterop.Release(mailMerge); }
    }

    private static object Merge(WordActionContext context, OfficeActionRequest request)
    {
        var dataSource = Path.GetFullPath(request.StringParam("dataSourcePath"));
        if (!File.Exists(dataSource)) throw new OfficeWorkerException("file_not_found", $"邮件合并数据源不存在: {dataSource}");
        var output = request.OutputPath ?? Path.Combine(Path.GetDirectoryName(request.FilePath) ?? Environment.CurrentDirectory, "merged.docx");
        output = Path.GetFullPath(output);
        var records = ReadSpreadsheetRecords(dataSource);
        if (records.Count == 0) throw new OfficeWorkerException("mail_merge_empty", "邮件合并数据源没有可用记录");
        var outputFormat = request.StringParam("outputFormat", Path.GetExtension(output).Equals(".pdf", StringComparison.OrdinalIgnoreCase) ? "pdf" : "docx").ToLowerInvariant();
        var files = new List<string>();

        if (request.Operation == "batchMailMerge")
        {
            var outputDirectory = request.StringParam("outputDirectory", Path.GetDirectoryName(output) ?? Environment.CurrentDirectory);
            Directory.CreateDirectory(outputDirectory);
            for (var index = 0; index < records.Count; index++)
            {
                var fileName = MergeFileName(request.StringParam("fileNamePattern", "document-{index}"), records[index], index + 1);
                using var merged = CreateMergedDocument(context, request, [records[index]]);
                SaveMergedDocument(merged.Document, Path.Combine(outputDirectory, fileName), outputFormat, files);
            }
            return OfficeActionResults.Done(request, "com", "已批量生成邮件合并文档", new { files, recordCount = records.Count }, [new OfficeChange("mail-merge", outputDirectory, "已批量生成邮件合并文档")], outputDirectory);
        }

        Directory.CreateDirectory(Path.GetDirectoryName(output) ?? Environment.CurrentDirectory);
        using (var merged = CreateMergedDocument(context, request, records))
        {
            SaveMergedDocument(merged.Document, output, outputFormat, files);
        }
        return OfficeActionResults.Done(request, "com", "已生成邮件合并文档", new { files, outputPath = files.FirstOrDefault() ?? output, recordCount = records.Count }, [new OfficeChange("mail-merge", output, "已生成邮件合并文档")], files.FirstOrDefault() ?? output);
    }

    private static MergedDocument CreateMergedDocument(WordActionContext context, OfficeActionRequest request, IReadOnlyList<Dictionary<string, string>> records)
    {
        object? documents = null; object? document = null;
        try
        {
            documents = context.App.Documents;
            document = ((dynamic)documents).Add();
            for (var index = 0; index < records.Count; index++) AppendRecord(context, document, request, records[index], index > 0);
            return new MergedDocument(document);
        }
        catch
        {
            if (document is not null) { try { ((dynamic)document).Close(0); } catch { } ComInterop.Release(document); }
            throw;
        }
        finally { ComInterop.Release(documents); }
    }

    private static void AppendRecord(WordActionContext context, object targetDocument, OfficeActionRequest request, Dictionary<string, string> record, bool pageBreak)
    {
        object? source = null; object? insertion = null; object? recordRange = null;
        try
        {
            dynamic target = targetDocument;
            var start = Math.Max(0, Convert.ToInt32(target.Content.End) - 1);
            insertion = target.Range(start, start); dynamic insertionApi = insertion;
            if (pageBreak) { insertionApi.InsertBreak(7); start = Math.Max(0, Convert.ToInt32(target.Content.End) - 1); insertion = target.Range(start, start); insertionApi = insertion; }
            source = context.Document.Content;
            insertionApi.FormattedText = ((dynamic)source).FormattedText;
            var end = Math.Max(start, Convert.ToInt32(target.Content.End) - 1);
            recordRange = target.Range(start, end);
            var values = new Dictionary<string, string>(record, StringComparer.OrdinalIgnoreCase);
            ApplyConditions(request.Param("conditions"), values);
            foreach (var pair in values) ReplaceAll(recordRange, $"{{{{{pair.Key}}}}}", pair.Value);
            ApplyImages(recordRange, request.Param("imageFields"), values);
        }
        finally { ComInterop.Release(recordRange); ComInterop.Release(insertion); ComInterop.Release(source); }
    }

    private static void ApplyConditions(JsonElement conditions, Dictionary<string, string> values)
    {
        if (conditions.ValueKind != JsonValueKind.Array) return;
        foreach (var condition in conditions.EnumerateArray())
        {
            var placeholder = ElementString(condition, "placeholder");
            var field = ElementString(condition, "field");
            if (placeholder.Length == 0 || field.Length == 0) continue;
            var actual = values.GetValueOrDefault(field, string.Empty);
            var expected = ElementString(condition, "value");
            var matches = ElementString(condition, "operator", "eq") switch
            {
                "ne" => !actual.Equals(expected, StringComparison.OrdinalIgnoreCase),
                "contains" => actual.Contains(expected, StringComparison.OrdinalIgnoreCase),
                _ => actual.Equals(expected, StringComparison.OrdinalIgnoreCase),
            };
            values[placeholder.Trim('{', '}')] = matches ? ElementString(condition, "trueText") : ElementString(condition, "falseText");
        }
    }

    private static void ApplyImages(object recordRange, JsonElement imageFields, IReadOnlyDictionary<string, string> values)
    {
        if (imageFields.ValueKind != JsonValueKind.Array) return;
        foreach (var imageField in imageFields.EnumerateArray())
        {
            var placeholder = ElementString(imageField, "placeholder");
            var path = values.GetValueOrDefault(ElementString(imageField, "field"), string.Empty);
            if (placeholder.Length == 0 || !File.Exists(path)) continue;
            object? searchRange = null; object? finder = null; object? images = null; object? image = null;
            try
            {
                searchRange = ((dynamic)recordRange).Duplicate; dynamic rangeApi = searchRange;
                finder = rangeApi.Find; dynamic findApi = finder; findApi.Text = placeholder; findApi.Wrap = 0;
                if (!Convert.ToBoolean(findApi.Execute())) continue;
                rangeApi.Text = string.Empty;
                images = rangeApi.InlineShapes;
                image = ((dynamic)images).AddPicture(path, false, true, searchRange);
                var width = imageField.TryGetProperty("width", out var widthValue) && widthValue.TryGetDouble(out var requestedWidth) ? requestedWidth : 0;
                if (width > 0) ((dynamic)image).Width = width;
            }
            finally { ComInterop.Release(image); ComInterop.Release(images); ComInterop.Release(finder); ComInterop.Release(searchRange); }
        }
    }

    private static void ReplaceAll(object range, string find, string replacementText)
    {
        object? finder = null; object? replacement = null;
        try
        {
            finder = ((dynamic)range).Find; dynamic findApi = finder;
            replacement = findApi.Replacement; ((dynamic)replacement).Text = replacementText;
            findApi.Text = find; findApi.Forward = true; findApi.Wrap = 0;
            _ = findApi.Execute(Replace: 2);
        }
        finally { ComInterop.Release(replacement); ComInterop.Release(finder); }
    }

    private static void SaveMergedDocument(object document, string output, string outputFormat, List<string> files)
    {
        var directory = Path.GetDirectoryName(output) ?? Environment.CurrentDirectory;
        Directory.CreateDirectory(directory);
        var basePath = Path.Combine(directory, Path.GetFileNameWithoutExtension(output));
        dynamic documentApi = document;
        if (outputFormat is "docx" or "both")
        {
            var docx = outputFormat == "docx" && Path.GetExtension(output).Equals(".docx", StringComparison.OrdinalIgnoreCase) ? output : basePath + ".docx";
            documentApi.SaveAs2(docx); files.Add(docx);
        }
        if (outputFormat is "pdf" or "both")
        {
            var pdf = outputFormat == "pdf" && Path.GetExtension(output).Equals(".pdf", StringComparison.OrdinalIgnoreCase) ? output : basePath + ".pdf";
            documentApi.ExportAsFixedFormat(pdf, 17); files.Add(pdf);
        }
    }

    private static string MergeFileName(string pattern, IReadOnlyDictionary<string, string> record, int index)
    {
        var name = pattern.Replace("{index}", index.ToString("D4"), StringComparison.OrdinalIgnoreCase);
        foreach (var pair in record) name = name.Replace($"{{{pair.Key}}}", pair.Value, StringComparison.OrdinalIgnoreCase);
        var invalid = Path.GetInvalidFileNameChars();
        name = new string(name.Select(character => invalid.Contains(character) ? '_' : character).ToArray()).Trim();
        return name.Length > 0 ? name : $"document-{index:D4}";
    }

    private static List<Dictionary<string, string>> ReadSpreadsheetRecords(string path)
    {
        using var document = SpreadsheetDocument.Open(path, false);
        var workbook = document.WorkbookPart ?? throw new OfficeWorkerException("invalid_file", "Excel 数据源缺少工作簿部件");
        var sheet = workbook.Workbook.Sheets?.Elements<S.Sheet>().FirstOrDefault() ?? throw new OfficeWorkerException("invalid_file", "Excel 数据源没有工作表");
        var worksheet = workbook.GetPartById(sheet.Id!.Value!) as WorksheetPart ?? throw new OfficeWorkerException("invalid_file", "Excel 数据源工作表关系无效");
        var rows = worksheet.Worksheet.GetFirstChild<S.SheetData>()?.Elements<S.Row>().ToArray() ?? [];
        if (rows.Length < 2) return [];
        var shared = workbook.SharedStringTablePart?.SharedStringTable?.Elements<S.SharedStringItem>().Select(item => item.InnerText).ToArray() ?? [];
        var headers = ReadSpreadsheetRow(rows[0], shared);
        return rows.Skip(1).Select(row =>
        {
            var values = ReadSpreadsheetRow(row, shared);
            return headers.Select((header, index) => new { header, value = index < values.Length ? values[index] : string.Empty })
                .Where(item => item.header.Length > 0).ToDictionary(item => item.header, item => item.value, StringComparer.OrdinalIgnoreCase);
        }).ToList();
    }

    private static string[] ReadSpreadsheetRow(S.Row row, IReadOnlyList<string> shared)
    {
        var cells = new SortedDictionary<int, string>();
        var next = 1;
        foreach (var cell in row.Elements<S.Cell>())
        {
            var column = SpreadsheetColumn(cell.CellReference?.Value);
            if (column == 0) column = next;
            var value = cell.CellValue?.Text ?? cell.InlineString?.InnerText ?? string.Empty;
            if (cell.DataType?.Value == S.CellValues.SharedString && int.TryParse(value, out var sharedIndex)) value = shared.ElementAtOrDefault(sharedIndex) ?? string.Empty;
            cells[column] = value; next = column + 1;
        }
        return Enumerable.Range(1, cells.Keys.DefaultIfEmpty(0).Max()).Select(index => cells.GetValueOrDefault(index, string.Empty)).ToArray();
    }

    private static int SpreadsheetColumn(string? reference)
    {
        var result = 0;
        foreach (var character in (reference ?? string.Empty).TakeWhile(char.IsLetter)) result = result * 26 + char.ToUpperInvariant(character) - 'A' + 1;
        return result;
    }

    private sealed class MergedDocument(object document) : IDisposable
    {
        public object Document { get; } = document;
        public void Dispose()
        {
            try { ((dynamic)Document).Close(0); } catch { }
            ComInterop.Release(Document);
        }
    }

    private static object? FindControl(dynamic collection, OfficeActionRequest request)
    {
        var id = request.StringParam("id"); var title = request.StringParam("title"); var tag = request.StringParam("tag");
        for (var index = 1; index <= Convert.ToInt32(collection.Count); index++)
        {
            object? control = collection.Item(index); dynamic api = control;
            if (id.Length > 0 && Convert.ToString(api.ID) == id || title.Length > 0 && Convert.ToString(api.Title) == title || tag.Length > 0 && Convert.ToString(api.Tag) == tag) return control;
            ComInterop.Release(control);
        }
        return null;
    }

    private static object AddControl(WordActionContext context, dynamic collection, JsonElement spec)
    {
        object? paragraphs = null; object? paragraph = null; object? range = null; object? control = null;
        try
        {
            paragraphs = context.Document.Paragraphs;
            paragraph = ((dynamic)paragraphs).Add();
            range = ((dynamic)paragraph).Range;
            dynamic rangeApi = range;
            rangeApi.End = Math.Max(Convert.ToInt32(rangeApi.Start), Convert.ToInt32(rangeApi.End) - 1);
            var requestedType = ElementString(spec, "type", "text");
            string? fallbackType = null;
            try { control = collection.Add(ControlType(requestedType), range); }
            catch (NotImplementedException)
            {
                control = collection.Add(ControlType("text"), range);
                fallbackType = requestedType.ToLowerInvariant();
            }
            ConfigureControl(control, spec, fallbackType);
            return new { type = ElementString(spec, "type", "text"), title = ElementString(spec, "title"), tag = ElementString(spec, "tag") };
        }
        finally { ComInterop.Release(control); ComInterop.Release(range); ComInterop.Release(paragraph); ComInterop.Release(paragraphs); }
    }

    private static void ConfigureControl(object control, JsonElement spec, string? fallbackType = null)
    {
        dynamic controlApi = control;
        var title = ElementString(spec, "title");
        var tag = ElementString(spec, "tag");
        var placeholder = ElementString(spec, "placeholder");
        if (title.Length > 0) controlApi.Title = title;
        if (tag.Length > 0) controlApi.Tag = fallbackType is null ? tag : $"wengge:{fallbackType}:{tag}";
        if (placeholder.Length > 0)
        {
            try { controlApi.SetPlaceholderText(null, null, placeholder); }
            catch (ArgumentException)
            {
                object? range = null;
                try { range = controlApi.Range; ((dynamic)range).Text = placeholder; }
                finally { ComInterop.Release(range); }
            }
        }
        if (spec.TryGetProperty("lockContents", out var lockContents) && lockContents.ValueKind is JsonValueKind.True or JsonValueKind.False) controlApi.LockContents = lockContents.GetBoolean();
        if (spec.TryGetProperty("lockControl", out var lockControl) && lockControl.ValueKind is JsonValueKind.True or JsonValueKind.False) controlApi.LockContentControl = lockControl.GetBoolean();
        if (spec.TryGetProperty("entries", out var entries) && entries.ValueKind == JsonValueKind.Array)
        {
            object? collection = null;
            try
            {
                collection = controlApi.DropdownListEntries; dynamic entriesApi = collection;
                foreach (var entry in entries.EnumerateArray())
                {
                    object? added = null;
                    try { added = entriesApi.Add(ElementString(entry, "text"), ElementString(entry, "value")); }
                    finally { ComInterop.Release(added); }
                }
            }
            finally { ComInterop.Release(collection); }
        }
        if (fallbackType == "checkbox")
        {
            object? range = null;
            try { range = controlApi.Range; ((dynamic)range).Text = "[ ]"; }
            finally { ComInterop.Release(range); }
        }
    }

    private static void SelectListEntry(dynamic control, string value)
    {
        object? entries = null;
        try
        {
            entries = control.DropdownListEntries; dynamic entriesApi = entries;
            for (var index = 1; index <= Convert.ToInt32(entriesApi.Count); index++)
            {
                object? entry = null;
                try
                {
                    entry = entriesApi.Item(index); dynamic entryApi = entry;
                    if (string.Equals(Convert.ToString(entryApi.Value), value, StringComparison.OrdinalIgnoreCase) || string.Equals(Convert.ToString(entryApi.Text), value, StringComparison.OrdinalIgnoreCase))
                    {
                        entryApi.Select();
                        return;
                    }
                }
                finally { ComInterop.Release(entry); }
            }
        }
        finally { ComInterop.Release(entries); }
    }

    private static string ElementValue(JsonElement value) => value.ValueKind == JsonValueKind.Object && value.TryGetProperty("value", out var nested)
        ? nested.ValueKind == JsonValueKind.String ? nested.GetString() ?? string.Empty : nested.ToString()
        : value.ValueKind == JsonValueKind.String ? value.GetString() ?? string.Empty : value.ToString();

    private static string ElementString(JsonElement value, string name, string fallback = "") =>
        value.ValueKind == JsonValueKind.Object && value.TryGetProperty(name, out var property) && property.ValueKind == JsonValueKind.String
            ? property.GetString() ?? fallback
            : fallback;

    private static bool BooleanValue(JsonElement value) => value.ValueKind == JsonValueKind.True
        || value.ValueKind == JsonValueKind.String && bool.TryParse(value.GetString(), out var result) && result;

    private static (string Tag, string? LogicalType) DecodeControlTag(string tag)
    {
        if (!tag.StartsWith("wengge:", StringComparison.OrdinalIgnoreCase)) return (tag, null);
        var separator = tag.IndexOf(':', "wengge:".Length);
        return separator < 0 ? (tag, null) : (tag[(separator + 1)..], tag["wengge:".Length..separator].ToLowerInvariant());
    }

    private static object ResolveRange(WordActionContext context, OfficeActionRequest request) => context.Document.Range(request.IntParam("start", Math.Max(0, Convert.ToInt32(context.Document.Content.End) - 1)), request.IntParam("end", Math.Max(0, Convert.ToInt32(context.Document.Content.End) - 1)));
    private static int ControlType(string type) => type.ToLowerInvariant() switch { "richtext" => 0, "text" => 1, "picture" => 2, "combobox" => 3, "dropdown" or "dropdownlist" => 4, "date" => 6, "checkbox" => 8, _ => 1 };
    private static object? Safe(Func<object?> value) { try { return value(); } catch { return null; } }
}
