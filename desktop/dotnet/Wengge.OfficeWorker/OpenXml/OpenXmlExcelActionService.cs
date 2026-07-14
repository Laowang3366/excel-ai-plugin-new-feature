using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using Wengge.OfficeWorker.Office;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.OpenXml;

internal sealed class OpenXmlExcelActionService(OpenXmlTableService tables)
{
    private static readonly HashSet<string> Operations =
    ["createWorkbook", "writeRange", "setDataValidation", "applyConditionalFormatting", "styleTable", "insertChart"];

    public static bool Supports(string operation) => Operations.Contains(operation);

    public object Execute(OfficeActionRequest request) => request.Operation switch
    {
        "createWorkbook" => CreateWorkbook(request),
        "writeRange" => WriteRange(request),
        "setDataValidation" => SetDataValidation(request),
        "applyConditionalFormatting" => ApplyConditionalFormatting(request),
        "styleTable" => StyleTable(request),
        "insertChart" => OfficeActionResults.NeedsCom(request, "Open XML 图表包生成尚未覆盖，需要 COM 执行"),
        _ => throw new OfficeWorkerException("unsupported_operation", $"不支持的 Excel Open XML 操作: {request.Operation}"),
    };

    private static object CreateWorkbook(OfficeActionRequest request)
    {
        var output = OutputPath(request, create: true);
        Directory.CreateDirectory(Path.GetDirectoryName(output) ?? Environment.CurrentDirectory);
        if (File.Exists(output)) File.Delete(output);
        var sheetNames = StringArray(request.Param("sheetNames"));
        if (sheetNames.Count == 0) sheetNames = ["Sheet1"];
        ValidateSheetNames(sheetNames);
        var values = Matrix(request.Param("values"));
        var (targetSheet, address) = request.ExcelTarget();
        if (string.IsNullOrWhiteSpace(targetSheet) || targetSheet == "Sheet1" && sheetNames[0] != "Sheet1") targetSheet = sheetNames[0];
        var start = ParseCell(FirstCell(request.StringParam("startCell", address)));

        using var document = SpreadsheetDocument.Create(output, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new S.Workbook();
        var sheets = workbookPart.Workbook.AppendChild(new S.Sheets());
        AddBaseStyles(workbookPart);
        for (var index = 0; index < sheetNames.Count; index++)
        {
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new S.Worksheet(new S.SheetData());
            sheets.Append(new S.Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = (uint)index + 1,
                Name = sheetNames[index],
            });
            if (string.Equals(sheetNames[index], targetSheet, StringComparison.OrdinalIgnoreCase) && values.Count > 0)
                WriteMatrix(worksheetPart.Worksheet.GetFirstChild<S.SheetData>()!, start, values, address);
            worksheetPart.Worksheet.Save();
        }
        workbookPart.Workbook.Save();
        return Done(request, "已使用 .NET Open XML 创建 Excel 工作簿", output,
            ["xl/workbook.xml", .. sheetNames.Select((_, index) => $"xl/worksheets/sheet{index + 1}.xml")],
            new { sheetNames, initialRange = values.Count > 0 ? request.Target : null, rowsWritten = values.Count });
    }

    private static object WriteRange(OfficeActionRequest request)
    {
        var values = Matrix(request.Param("values"));
        if (values.Count == 0) throw new OfficeWorkerException("invalid_params", "writeRange 操作需要 params.values 二维数组");
        var output = PrepareCopy(request, "advanced");
        using var document = SpreadsheetDocument.Open(output, true);
        var worksheet = ResolveWorksheet(document, request, out var partName);
        var (_, address) = request.ExcelTarget();
        WriteMatrix(worksheet.Worksheet.GetFirstChild<S.SheetData>() ?? worksheet.Worksheet.AppendChild(new S.SheetData()), ParseCell(FirstCell(address)), values, address);
        worksheet.Worksheet.Save();
        return Done(request, "已使用 .NET Open XML 写入 Excel 单元格", output, [partName],
            new { rowsWritten = values.Count, columnsWritten = values.Max(row => row.Count) });
    }

    private static object SetDataValidation(OfficeActionRequest request)
    {
        var output = PrepareCopy(request, "advanced", defaultToSource: false);
        using var document = SpreadsheetDocument.Open(output, true);
        var worksheetPart = ResolveWorksheet(document, request, out var partName);
        var (_, address) = request.ExcelTarget();
        var values = StringArray(request.Param("values"));
        var type = request.StringParam("type", "list");
        var validations = worksheetPart.Worksheet.Elements<S.DataValidations>().FirstOrDefault();
        if (validations is null)
        {
            validations = new S.DataValidations();
            InsertWorksheetChild(worksheetPart.Worksheet, validations);
        }
        validations.Append(new S.DataValidation
        {
            Type = ValidationType(type),
            AllowBlank = true,
            SequenceOfReferences = new ListValue<StringValue> { InnerText = address },
            Formula1 = new S.Formula1($"\"{string.Join(',', values)}\""),
        });
        validations.Count = (uint)validations.Elements<S.DataValidation>().Count();
        worksheetPart.Worksheet.Save();
        return Done(request, "已写入 Excel 数据验证", output, [partName]);
    }

    private static object ApplyConditionalFormatting(OfficeActionRequest request)
    {
        var output = PrepareCopy(request, "advanced", defaultToSource: false);
        using var document = SpreadsheetDocument.Open(output, true);
        var worksheetPart = ResolveWorksheet(document, request, out var partName);
        var workbookPart = document.WorkbookPart ?? throw new OfficeWorkerException("invalid_file", "Excel 工作簿缺少 workbook 部件");
        var styles = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
        styles.Stylesheet ??= BaseStylesheet();
        styles.Stylesheet.DifferentialFormats ??= new S.DifferentialFormats();
        var color = NormalizeColor(request.StringParam("fillColor", "FFF2CC"));
        var dxfId = (uint)styles.Stylesheet.DifferentialFormats.ChildElements.Count;
        styles.Stylesheet.DifferentialFormats.Append(new S.DifferentialFormat(
            new S.Fill(new S.PatternFill(
                new S.ForegroundColor { Rgb = $"FF{color}" },
                new S.BackgroundColor { Indexed = 64U }) { PatternType = S.PatternValues.Solid })));
        styles.Stylesheet.DifferentialFormats.Count = (uint)styles.Stylesheet.DifferentialFormats.ChildElements.Count;
        styles.Stylesheet.Save();
        var (_, address) = request.ExcelTarget();
        var priority = worksheetPart.Worksheet.Descendants<S.ConditionalFormattingRule>().Select(rule => rule.Priority?.Value ?? 0).DefaultIfEmpty().Max() + 1;
        var formatting = new S.ConditionalFormatting(
            new S.ConditionalFormattingRule(new S.Formula("TRUE"))
            {
                Type = S.ConditionalFormatValues.Expression,
                Priority = priority,
                FormatId = dxfId,
            })
        {
            SequenceOfReferences = new ListValue<StringValue> { InnerText = address },
        };
        InsertWorksheetChild(worksheetPart.Worksheet, formatting);
        worksheetPart.Worksheet.Save();
        return Done(request, "已写入 Excel 条件格式", output, [partName, "xl/styles.xml"]);
    }

    private object StyleTable(OfficeActionRequest request)
    {
        var source = RequireFile(request.FilePath);
        var data = tables.ApplyStyle(source, request.StringParam("style", "professional"), request.OutputPath, request.Target);
        var output = ReadProperty(data, "outputPath") ?? request.OutputPath ?? source;
        return OfficeActionResults.Done(request, "openxml", "已应用 Excel 表格样式", data,
            [new OfficeChange("openxml-part", request.Target, "已更新 Excel 表格样式")], output);
    }

    private static WorksheetPart ResolveWorksheet(SpreadsheetDocument document, OfficeActionRequest request, out string partName)
    {
        var workbook = document.WorkbookPart ?? throw new OfficeWorkerException("invalid_file", "Excel 工作簿缺少 workbook 部件");
        var (sheetName, _) = request.ExcelTarget();
        var sheet = workbook.Workbook.Sheets?.Elements<S.Sheet>().FirstOrDefault(candidate => string.Equals(candidate.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase))
            ?? workbook.Workbook.Sheets?.Elements<S.Sheet>().FirstOrDefault()
            ?? throw new OfficeWorkerException("sheet_not_found", $"找不到工作表: {sheetName}");
        var part = (WorksheetPart)workbook.GetPartById(sheet.Id!);
        partName = part.Uri.ToString().TrimStart('/');
        return part;
    }

    private static void WriteMatrix(S.SheetData sheetData, CellAddress start, IReadOnlyList<IReadOnlyList<JsonElement>> values, string targetAddress)
    {
        var hasDynamic = values.SelectMany(row => row).Any(IsDynamicFormula);
        var targetRef = hasDynamic && targetAddress.Contains(':') ? targetAddress : null;
        for (var rowOffset = 0; rowOffset < values.Count; rowOffset++)
        {
            var rowIndex = (uint)(start.Row + rowOffset);
            var row = sheetData.Elements<S.Row>().FirstOrDefault(candidate => candidate.RowIndex?.Value == rowIndex);
            if (row is null)
            {
                row = new S.Row { RowIndex = rowIndex };
                var next = sheetData.Elements<S.Row>().FirstOrDefault(candidate => candidate.RowIndex?.Value > rowIndex);
                if (next is null) sheetData.Append(row); else sheetData.InsertBefore(row, next);
            }
            for (var columnOffset = 0; columnOffset < values[rowOffset].Count; columnOffset++)
            {
                var reference = $"{ColumnName(start.Column + columnOffset)}{rowIndex}";
                var existing = row.Elements<S.Cell>().FirstOrDefault(cell => string.Equals(cell.CellReference?.Value, reference, StringComparison.OrdinalIgnoreCase));
                existing?.Remove();
                var value = values[rowOffset][columnOffset];
                if (value.ValueKind == JsonValueKind.String && value.GetString() == string.Empty && hasDynamic) continue;
                var cell = BuildCell(reference, value, targetRef);
                var nextCell = row.Elements<S.Cell>().FirstOrDefault(candidate => ColumnIndex(candidate.CellReference?.Value) > start.Column + columnOffset);
                if (nextCell is null) row.Append(cell); else row.InsertBefore(cell, nextCell);
            }
            if (!row.Elements<S.Cell>().Any() && !row.ChildElements.Any(child => child is not S.Cell)) row.Remove();
        }
    }

    private static S.Cell BuildCell(string reference, JsonElement value, string? targetRef)
    {
        var cell = new S.Cell { CellReference = reference };
        switch (value.ValueKind)
        {
            case JsonValueKind.Number:
                cell.CellValue = new S.CellValue(value.GetRawText());
                cell.DataType = S.CellValues.Number;
                break;
            case JsonValueKind.True:
            case JsonValueKind.False:
                cell.CellValue = new S.CellValue(value.GetBoolean() ? "1" : "0");
                cell.DataType = S.CellValues.Boolean;
                break;
            case JsonValueKind.String:
                var text = value.GetString() ?? string.Empty;
                if (text.StartsWith('=') && text.Length > 1)
                {
                    var formula = NormalizeFormula(text[1..]);
                    cell.CellFormula = new S.CellFormula(formula);
                    if (IsDynamicFormula(value) && !string.IsNullOrWhiteSpace(targetRef))
                    {
                        cell.CellFormula.FormulaType = S.CellFormulaValues.Array;
                        cell.CellFormula.Reference = targetRef;
                    }
                }
                else
                {
                    cell.DataType = S.CellValues.InlineString;
                    cell.InlineString = new S.InlineString(new S.Text(text) { Space = SpaceProcessingModeValues.Preserve });
                }
                break;
            default:
                cell.DataType = S.CellValues.InlineString;
                cell.InlineString = new S.InlineString(new S.Text(value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined ? string.Empty : value.ToString()));
                break;
        }
        return cell;
    }

    private static bool IsDynamicFormula(JsonElement value)
    {
        if (value.ValueKind != JsonValueKind.String) return false;
        var formula = value.GetString() ?? string.Empty;
        return formula.StartsWith('=') && new[] { "FILTER(", "SORT(", "SORTBY(", "UNIQUE(", "SEQUENCE(", "RANDARRAY(", "TOCOL(", "TOROW(", "WRAPROWS(", "WRAPCOLS(" }
            .Any(name => formula.Contains(name, StringComparison.OrdinalIgnoreCase));
    }

    private static string NormalizeFormula(string formula)
    {
        var dynamicFunctions = new[] { "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "RANDARRAY", "TOCOL", "TOROW", "WRAPROWS", "WRAPCOLS" };
        foreach (var function in dynamicFunctions)
            if (formula.StartsWith($"{function}(", StringComparison.OrdinalIgnoreCase)) return $"_xlfn._xlws.{formula}";
        return formula;
    }

    private static List<IReadOnlyList<JsonElement>> Matrix(JsonElement value)
    {
        if (value.ValueKind != JsonValueKind.Array) return [];
        var items = value.EnumerateArray().ToArray();
        if (items.Length == 0) return [];
        return items.All(item => item.ValueKind == JsonValueKind.Array)
            ? items.Select(item => (IReadOnlyList<JsonElement>)item.EnumerateArray().ToArray()).ToList()
            : [(IReadOnlyList<JsonElement>)items];
    }

    private static List<string> StringArray(JsonElement value) => value.ValueKind == JsonValueKind.Array
        ? value.EnumerateArray().Where(item => item.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(item.GetString())).Select(item => item.GetString()!.Trim()).ToList()
        : [];

    private static void ValidateSheetNames(IEnumerable<string> names)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var name in names)
        {
            if (name.Length > 31 || name.IndexOfAny(['[', ']', ':', '*', '?', '/', '\\']) >= 0)
                throw new OfficeWorkerException("invalid_params", $"工作表名称无效: {name}");
            if (!seen.Add(name)) throw new OfficeWorkerException("invalid_params", $"工作表名称重复: {name}");
        }
    }

    private static string PrepareCopy(OfficeActionRequest request, string suffix, bool defaultToSource = true)
    {
        var source = RequireFile(request.FilePath);
        var output = !string.IsNullOrWhiteSpace(request.OutputPath) ? Path.GetFullPath(request.OutputPath)
            : defaultToSource ? source : DefaultOutputPath(source, suffix);
        Directory.CreateDirectory(Path.GetDirectoryName(output) ?? Environment.CurrentDirectory);
        if (!string.Equals(source, output, StringComparison.OrdinalIgnoreCase)) File.Copy(source, output, true);
        return output;
    }

    private static string OutputPath(OfficeActionRequest request, bool create) =>
        Path.GetFullPath(request.OutputPath ?? request.FilePath ?? (create ? throw new OfficeWorkerException("invalid_params", "createWorkbook 需要 filePath 或 outputPath") : string.Empty));

    private static string RequireFile(string? path)
    {
        if (string.IsNullOrWhiteSpace(path)) throw new OfficeWorkerException("invalid_params", "缺少 filePath");
        var fullPath = Path.GetFullPath(path);
        return File.Exists(fullPath) ? fullPath : throw new OfficeWorkerException("file_not_found", $"Office 文件不存在: {fullPath}");
    }

    private static object Done(OfficeActionRequest request, string summary, string output, IEnumerable<string> parts, object? data = null) =>
        OfficeActionResults.Done(request, "openxml", summary, data ?? new { outputPath = output, changedParts = parts.ToArray() },
            parts.Select(part => new OfficeChange("openxml-part", part, $"已更新 {part}")), output);

    private static void AddBaseStyles(WorkbookPart workbookPart)
    {
        var styles = workbookPart.AddNewPart<WorkbookStylesPart>();
        styles.Stylesheet = BaseStylesheet();
        styles.Stylesheet.Save();
    }

    private static S.Stylesheet BaseStylesheet() => new(
        new S.Fonts(new S.Font()) { Count = 1 },
        new S.Fills(new S.Fill(new S.PatternFill { PatternType = S.PatternValues.None }), new S.Fill(new S.PatternFill { PatternType = S.PatternValues.Gray125 })) { Count = 2 },
        new S.Borders(new S.Border()) { Count = 1 },
        new S.CellStyleFormats(new S.CellFormat()) { Count = 1 },
        new S.CellFormats(new S.CellFormat()) { Count = 1 },
        new S.CellStyles(new S.CellStyle { Name = "Normal", FormatId = 0U, BuiltinId = 0U }) { Count = 1 });

    private static S.DataValidationValues ValidationType(string type) => type.ToLowerInvariant() switch
    {
        "whole" => S.DataValidationValues.Whole,
        "decimal" => S.DataValidationValues.Decimal,
        "date" => S.DataValidationValues.Date,
        "time" => S.DataValidationValues.Time,
        "textlength" => S.DataValidationValues.TextLength,
        "custom" => S.DataValidationValues.Custom,
        _ => S.DataValidationValues.List,
    };

    private static void InsertWorksheetChild(S.Worksheet worksheet, OpenXmlElement element)
    {
        var order = WorksheetChildOrder(element.LocalName);
        var next = worksheet.ChildElements.FirstOrDefault(child => WorksheetChildOrder(child.LocalName) > order);
        if (next is null) worksheet.Append(element); else worksheet.InsertBefore(element, next);
    }

    private static int WorksheetChildOrder(string localName) => localName switch
    {
        "sheetPr" => 0, "dimension" => 1, "sheetViews" => 2, "sheetFormatPr" => 3, "cols" => 4,
        "sheetData" => 5, "sheetCalcPr" => 6, "sheetProtection" => 7, "protectedRanges" => 8,
        "scenarios" => 9, "autoFilter" => 10, "sortState" => 11, "dataConsolidate" => 12,
        "customSheetViews" => 13, "mergeCells" => 14, "phoneticPr" => 15,
        "conditionalFormatting" => 16, "dataValidations" => 17, "hyperlinks" => 18,
        "printOptions" => 19, "pageMargins" => 20, "pageSetup" => 21, "headerFooter" => 22,
        "rowBreaks" => 23, "colBreaks" => 24, "customProperties" => 25, "cellWatches" => 26,
        "ignoredErrors" => 27, "smartTags" => 28, "drawing" => 29, "legacyDrawing" => 30,
        "legacyDrawingHF" => 31, "picture" => 32, "oleObjects" => 33, "controls" => 34,
        "webPublishItems" => 35, "tableParts" => 36, "extLst" => 37, _ => int.MaxValue,
    };

    private static string NormalizeColor(string color)
    {
        var normalized = color.Trim().TrimStart('#');
        return normalized.Length == 8 ? normalized[2..].ToUpperInvariant()
            : normalized.Length == 6 && normalized.All(Uri.IsHexDigit) ? normalized.ToUpperInvariant() : "FFF2CC";
    }

    private static string? ReadProperty(object value, string property) => value.GetType().GetProperty(property)?.GetValue(value) as string;
    private static string DefaultOutputPath(string path, string suffix) => Path.Combine(Path.GetDirectoryName(path) ?? Environment.CurrentDirectory, $"{Path.GetFileNameWithoutExtension(path)}-{suffix}{Path.GetExtension(path)}");
    private static string FirstCell(string address) => address.Split(':', 2)[0].Trim();

    private static CellAddress ParseCell(string address)
    {
        var letters = new string(address.TakeWhile(char.IsLetter).ToArray());
        var digits = new string(address.SkipWhile(char.IsLetter).TakeWhile(char.IsDigit).ToArray());
        if (letters.Length == 0 || !int.TryParse(digits, out var row) || row < 1) throw new OfficeWorkerException("invalid_params", $"单元格地址无效: {address}");
        var column = 0;
        foreach (var letter in letters.ToUpperInvariant()) column = column * 26 + letter - 'A' + 1;
        return new CellAddress(column, row);
    }

    private static string ColumnName(int index)
    {
        var result = string.Empty;
        while (index > 0) { index--; result = (char)('A' + index % 26) + result; index /= 26; }
        return result;
    }

    private static int ColumnIndex(string? reference)
    {
        if (string.IsNullOrWhiteSpace(reference)) return int.MaxValue;
        var result = 0;
        foreach (var letter in reference.TakeWhile(char.IsLetter)) result = result * 26 + char.ToUpperInvariant(letter) - 'A' + 1;
        return result;
    }

    private sealed record CellAddress(int Column, int Row);
}
