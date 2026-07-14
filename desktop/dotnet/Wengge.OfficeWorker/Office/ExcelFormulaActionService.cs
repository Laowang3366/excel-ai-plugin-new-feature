using System.Text.Json;
using System.Text.RegularExpressions;
using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Excel;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Office;

internal sealed partial class ExcelFormulaActionService(OfficeApplicationProvider applications)
{
    private static readonly HashSet<string> Operations =
    [
        "traceFormulaDependencies", "inspectFormulaDependencies", "repairFormulaReferences",
        "convertFormulasToValues", "inspectFormulaBackups", "restoreFormulas",
        "inspectFormulaProtection", "manageFormulaProtection",
    ];

    public static bool Supports(string operation) => Operations.Contains(operation);

    public object Execute(OfficeActionRequest request)
    {
        using var context = new ExcelActionContext(applications, request);
        var readOnly = request.Operation is "traceFormulaDependencies" or "inspectFormulaDependencies" or "inspectFormulaBackups" or "inspectFormulaProtection";
        var data = request.Operation switch
        {
            "traceFormulaDependencies" or "inspectFormulaDependencies" => DependencyReport(context, request),
            "repairFormulaReferences" => RepairReferences(context, request),
            "convertFormulasToValues" => ConvertToValues(context, request),
            "inspectFormulaBackups" => InspectBackups(context),
            "restoreFormulas" => Restore(context, request),
            "inspectFormulaProtection" => InspectProtection(context, request),
            "manageFormulaProtection" => ManageProtection(context, request),
            _ => throw new OfficeWorkerException("unsupported_operation", $"不支持的公式治理操作: {request.Operation}"),
        };
        if (!readOnly) context.Save(request);
        return OfficeActionResults.Done(request, "com", FormulaSummary(request.Operation), data,
            readOnly ? [] : [new OfficeChange("formula-governance", request.Target, FormulaSummary(request.Operation))]);
    }

    private static object DependencyReport(ExcelActionContext context, OfficeActionRequest request)
    {
        var nodes = ReadFormulaNodes(context, request);
        var edges = new List<FormulaEdge>();
        var broken = new List<object>();
        foreach (var node in nodes)
        {
            if (node.Formula.Contains("#REF!", StringComparison.OrdinalIgnoreCase)) broken.Add(new { cell = node.Id, formula = node.Formula, reason = "#REF!" });
            foreach (Match match in QualifiedReferenceRegex().Matches(RemoveStringLiterals(node.Formula)))
            {
                var sheet = (match.Groups["quoted"].Success ? match.Groups["quoted"].Value.Replace("''", "'", StringComparison.Ordinal) : match.Groups["plain"].Value).Trim();
                var address = match.Groups["address"].Value.Replace("$", string.Empty, StringComparison.Ordinal);
                edges.Add(new FormulaEdge(node.Id, $"{sheet}!{address}", "cross-sheet", match.Value));
            }
            var localFormula = QualifiedReferenceRegex().Replace(RemoveStringLiterals(node.Formula), string.Empty);
            foreach (Match match in LocalReferenceRegex().Matches(localFormula))
            {
                var address = match.Groups["address"].Value.Replace("$", string.Empty, StringComparison.Ordinal);
                edges.Add(new FormulaEdge(node.Id, $"{node.Sheet}!{address}", "same-sheet", match.Value));
            }
            foreach (Match match in ExternalReferenceRegex().Matches(node.Formula))
                edges.Add(new FormulaEdge(node.Id, $"external:{match.Value}", "external", match.Value));
        }
        edges = edges.DistinctBy(edge => (edge.From, edge.To, edge.Kind)).ToList();
        var nodeIds = nodes.Select(node => node.Id).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var precedents = edges.GroupBy(edge => edge.From, StringComparer.OrdinalIgnoreCase).ToDictionary(group => group.Key, group => group.Select(edge => edge.To).Distinct(StringComparer.OrdinalIgnoreCase).ToArray(), StringComparer.OrdinalIgnoreCase);
        var dependents = edges.Where(edge => nodeIds.Contains(edge.To)).GroupBy(edge => edge.To, StringComparer.OrdinalIgnoreCase).ToDictionary(group => group.Key, group => group.Select(edge => edge.From).Distinct(StringComparer.OrdinalIgnoreCase).ToArray(), StringComparer.OrdinalIgnoreCase);
        var outputNodes = nodes.Select(node => new
        {
            id = node.Id,
            sheet = node.Sheet,
            address = node.Address,
            formula = node.Formula,
            value = node.Value,
            precedents = precedents.GetValueOrDefault(node.Id, []),
            dependents = dependents.GetValueOrDefault(node.Id, []),
        }).ToArray();
        var cycles = FindCycles(nodeIds, edges).Select(path => new { path }).ToArray();
        string? circularReference = null;
        try
        {
            object? cell = context.App.CircularReference;
            if (cell is not null)
            {
                try { circularReference = CellId(cell); }
                finally { ComInterop.Release(cell); }
            }
        }
        catch { }
        return new { nodes = outputNodes, edges = edges.Select(edge => new { from = edge.From, to = edge.To, kind = edge.Kind, reference = edge.Reference }), cycles, circularReference, brokenReferences = broken, formulaCount = nodes.Count, edgeCount = edges.Count };
    }

    private static object RepairReferences(ExcelActionContext context, OfficeActionRequest request)
    {
        var repairs = new List<object>();
        var unresolved = new List<object>();
        foreach (var cell in FormulaCells(context, request))
        {
            try
            {
                dynamic cellApi = cell;
                var before = Formula(cellApi);
                if (!before.Contains("#REF!", StringComparison.OrdinalIgnoreCase) && !request.BoolParam("applyAllMappings")) continue;
                var after = before;
                var replacements = request.Param("replacements");
                if (replacements.ValueKind == JsonValueKind.Array)
                {
                    foreach (var replacement in replacements.EnumerateArray())
                    {
                        if (!replacement.TryGetProperty("find", out var find) || find.ValueKind != JsonValueKind.String) continue;
                        var replace = replacement.TryGetProperty("replace", out var replaceValue) ? replaceValue.GetString() ?? string.Empty : string.Empty;
                        after = after.Replace(find.GetString()!, replace, StringComparison.Ordinal);
                    }
                }
                if (!string.Equals(after, before, StringComparison.Ordinal)) SetFormula(cellApi, after);
                var current = Formula(cellApi);
                var item = new { cell = CellId(cell), before, after = current, strategy = "mapping" };
                if (current.Contains("#REF!", StringComparison.OrdinalIgnoreCase)) unresolved.Add(item); else repairs.Add(item);
            }
            finally { ComInterop.Release(cell); }
        }
        try { context.App.Calculate(); } catch { }
        if (unresolved.Count > 0)
            throw new OfficeWorkerException("formula_repair_incomplete", "仍有公式引用未修复，未保存本次修改", new { repairs, unresolved });
        return new { repairs, repairedCount = repairs.Count, unresolved, unresolvedCount = unresolved.Count };
    }

    private static object ConvertToValues(ExcelActionContext context, OfficeActionRequest request)
    {
        var backup = request.BoolParam("createBackup", true);
        var backupId = request.StringParam("backupId", Guid.NewGuid().ToString("N"));
        object? backupSheet = backup ? GetBackupSheet(context.Workbook, create: true) : null;
        var row = backupSheet is null ? 0 : LastBackupRow(backupSheet) + 1;
        var converted = 0;
        try
        {
            foreach (var cell in FormulaCells(context, request))
            {
                try
                {
                    dynamic cellApi = cell;
                    if (backupSheet is not null) WriteBackup(backupSheet, row++, backupId, cellApi, request.Target ?? request.ExcelTarget().Address);
                    var value = cellApi.Value2;
                    cellApi.Value2 = value;
                    converted++;
                }
                finally { ComInterop.Release(cell); }
            }
            if (backupSheet is not null) ((dynamic)backupSheet).Visible = 2;
            return new { backupId = backup ? backupId : null, convertedFormulaCells = converted };
        }
        finally { ComInterop.Release(backupSheet); }
    }

    private static object InspectBackups(ExcelActionContext context)
    {
        object? sheet = GetBackupSheet(context.Workbook, create: false);
        try
        {
            var backups = sheet is null ? [] : BackupSummaries(sheet).ToArray();
            return new { backups, backupCount = backups.Length };
        }
        finally { ComInterop.Release(sheet); }
    }

    private static object Restore(ExcelActionContext context, OfficeActionRequest request)
    {
        object? backupSheet = GetBackupSheet(context.Workbook, create: false);
        if (backupSheet is null) throw new OfficeWorkerException("formula_backup_not_found", "当前工作簿没有公式备份");
        try
        {
            var summaries = BackupSummaries(backupSheet).ToArray();
            var backupId = request.StringParam("backupId", summaries.FirstOrDefault()?.BackupId ?? string.Empty);
            if (backupId.Length == 0) throw new OfficeWorkerException("formula_backup_not_found", "找不到可恢复的公式备份");
            var restored = new List<object>();
            var failed = new List<object>();
            var rows = new List<int>();
            dynamic sheetApi = backupSheet;
            for (var row = 3; row <= LastBackupRow(backupSheet); row++)
            {
                if (!string.Equals(Convert.ToString(sheetApi.Cells.Item(row, 1).Value2), backupId, StringComparison.Ordinal)) continue;
                rows.Add(row);
                var sheetName = Convert.ToString(sheetApi.Cells.Item(row, 3).Value2) ?? string.Empty;
                var address = Convert.ToString(sheetApi.Cells.Item(row, 4).Value2) ?? string.Empty;
                object? targetSheet = null;
                object? cell = null;
                try
                {
                    targetSheet = context.Workbook.Worksheets.Item(sheetName);
                    cell = ((dynamic)targetSheet).Range(address);
                    dynamic cellApi = cell;
                    var formula = Convert.ToString(sheetApi.Cells.Item(row, 5).Value2) ?? string.Empty;
                    SetFormula(cellApi, formula);
                    cellApi.NumberFormat = Convert.ToString(sheetApi.Cells.Item(row, 7).Value2);
                    cellApi.Locked = Convert.ToString(sheetApi.Cells.Item(row, 8).Value2) == "1";
                    restored.Add(new { cell = $"{sheetName}!{address}", formula });
                }
                catch (Exception exception) { failed.Add(new { cell = $"{sheetName}!{address}", error = exception.Message }); }
                finally { ComInterop.Release(cell); ComInterop.Release(targetSheet); }
            }
            if (request.BoolParam("removeAfterRestore")) foreach (var row in rows.OrderDescending()) sheetApi.Rows.Item(row).Delete();
            sheetApi.Visible = 2;
            try { context.App.Calculate(); } catch { }
            if (failed.Count > 0)
                throw new OfficeWorkerException("formula_restore_incomplete", "一个或多个公式未能恢复，未保存本次修改", new { backupId, restored, failed });
            return new { backupId, restored, restoredCount = restored.Count, failed };
        }
        finally { ComInterop.Release(backupSheet); }
    }

    private static ProtectionInspection InspectProtection(ExcelActionContext context, OfficeActionRequest request)
    {
        var items = new List<object>();
        foreach (var scope in Scopes(context, request))
        {
            var count = 0; var locked = 0;
            foreach (var cell in FormulaCells(scope.Range))
            {
                try { count++; if (Convert.ToBoolean(((dynamic)cell).Locked)) locked++; }
                finally { ComInterop.Release(cell); }
            }
            items.Add(new { sheet = Convert.ToString(((dynamic)scope.Sheet).Name), @protected = Convert.ToBoolean(((dynamic)scope.Sheet).ProtectContents), target = ((dynamic)scope.Range).Address[false, false], formulaCount = count, lockedFormulaCount = locked });
            scope.Dispose();
        }
        return new ProtectionInspection(items);
    }

    private static object ManageProtection(ExcelActionContext context, OfficeActionRequest request)
    {
        var command = request.StringParam("command", "lock");
        var password = request.StringParam("password");
        foreach (var scope in Scopes(context, request))
        {
            try
            {
                dynamic sheet = scope.Sheet; dynamic range = scope.Range;
                try { sheet.Unprotect(password); } catch { }
                if (command == "lock")
                {
                    if (request.BoolParam("unlockInputs", true)) range.Locked = false;
                    foreach (var cell in FormulaCells(scope.Range)) { try { ((dynamic)cell).Locked = true; } finally { ComInterop.Release(cell); } }
                    if (request.BoolParam("protectSheet", true)) sheet.Protect(password, true, true, true, true);
                }
                else if (command == "unlock")
                    foreach (var cell in FormulaCells(scope.Range)) { try { ((dynamic)cell).Locked = false; } finally { ComInterop.Release(cell); } }
                else throw new OfficeWorkerException("unsupported_operation", $"不支持的公式保护命令: {command}");
            }
            finally { scope.Dispose(); }
        }
        return new { command, protection = InspectProtection(context, request).Protection };
    }

    private static List<FormulaNode> ReadFormulaNodes(ExcelActionContext context, OfficeActionRequest request)
    {
        var nodes = new List<FormulaNode>();
        foreach (var cell in FormulaCells(context, request))
        {
            try
            {
                dynamic cellApi = cell;
                object? worksheet = null;
                try
                {
                    worksheet = cellApi.Worksheet; dynamic worksheetApi = worksheet;
                    nodes.Add(new FormulaNode(CellId(cell), Convert.ToString(worksheetApi.Name) ?? string.Empty, Convert.ToString(cellApi.Address[false, false]) ?? string.Empty, Formula(cellApi), cellApi.Value2));
                }
                finally { ComInterop.Release(worksheet); }
            }
            finally { ComInterop.Release(cell); }
        }
        return nodes;
    }

    private static IEnumerable<object> FormulaCells(ExcelActionContext context, OfficeActionRequest request)
    {
        var cells = new List<object>();
        foreach (var scope in Scopes(context, request))
        {
            try { cells.AddRange(FormulaCells(scope.Range)); }
            finally { scope.Dispose(); }
        }
        return cells;
    }

    private static IEnumerable<object> FormulaCells(object range)
    {
        object? formulaRange = null;
        object? areas = null;
        var result = new List<object>();
        try
        {
            try { formulaRange = ((dynamic)range).SpecialCells(-4123); }
            catch
            {
                object? cells = null;
                try
                {
                    cells = ((dynamic)range).Cells;
                    dynamic cellsApi = cells;
                    for (var cellIndex = 1; cellIndex <= Convert.ToInt32(cellsApi.Count); cellIndex++)
                    {
                        object? cell = null;
                        try
                        {
                            cell = cellsApi.Item(cellIndex);
                            if (Convert.ToBoolean(((dynamic)cell).HasFormula))
                            {
                                result.Add(cell);
                                cell = null;
                            }
                        }
                        catch
                        {
                            // Continue past cells that do not expose formula metadata.
                        }
                        finally { ComInterop.Release(cell); }
                    }
                    return result;
                }
                finally { ComInterop.Release(cells); }
            }
            areas = ((dynamic)formulaRange).Areas;
            dynamic areasApi = areas;
            for (var areaIndex = 1; areaIndex <= Convert.ToInt32(areasApi.Count); areaIndex++)
            {
                object? area = null;
                object? cells = null;
                try
                {
                    area = areasApi.Item(areaIndex);
                    cells = ((dynamic)area).Cells;
                    dynamic cellsApi = cells;
                    for (var cellIndex = 1; cellIndex <= Convert.ToInt32(cellsApi.Count); cellIndex++) result.Add(cellsApi.Item(cellIndex));
                }
                finally
                {
                    ComInterop.Release(cells);
                    ComInterop.Release(area);
                }
            }
            return result;
        }
        finally { ComInterop.Release(areas); ComInterop.Release(formulaRange); }
    }

    private static IEnumerable<FormulaScope> Scopes(ExcelActionContext context, OfficeActionRequest request)
    {
        var scope = request.StringParam("scope", request.Target is null ? "workbook" : "target");
        var output = new List<FormulaScope>();
        if (scope == "workbook")
        {
            object? sheets = null;
            try
            {
                sheets = context.Workbook.Worksheets; dynamic sheetsApi = sheets;
                for (var index = 1; index <= Convert.ToInt32(sheetsApi.Count); index++)
                {
                    object sheet = sheetsApi.Item(index); object range = ((dynamic)sheet).UsedRange;
                    output.Add(new FormulaScope(sheet, range));
                }
            }
            finally { ComInterop.Release(sheets); }
        }
        else
        {
            var (sheet, range) = context.GetRange(scope == "sheet" ? request with { Target = null } : request);
            if (scope == "sheet") { ComInterop.Release(range); range = ((dynamic)sheet).UsedRange; }
            output.Add(new FormulaScope(sheet, range));
        }
        return output;
    }

    private static object? GetBackupSheet(dynamic workbook, bool create)
    {
        object? sheets = null;
        try
        {
            sheets = workbook.Worksheets; dynamic sheetsApi = sheets;
            for (var index = 1; index <= Convert.ToInt32(sheetsApi.Count); index++)
            {
                object? sheet = sheetsApi.Item(index);
                try
                {
                    dynamic sheetApi = sheet!;
                    if ((Convert.ToString(sheetApi.Name) ?? string.Empty).StartsWith("_WenggeFormulaBackup", StringComparison.Ordinal) && Convert.ToString(sheetApi.Range("A1").Value2) == "WENGGE_FORMULA_BACKUP_V1") return sheet;
                }
                catch { }
                ComInterop.Release(sheet);
            }
            if (!create) return null;
            object backup = sheetsApi.Add(After: sheetsApi.Item(sheetsApi.Count)); dynamic backupApi = backup;
            backupApi.Name = $"_WenggeFormulaBackup{DateTime.Now:HHmmss}";
            backupApi.Range("A1").Value2 = "WENGGE_FORMULA_BACKUP_V1";
            var headers = new[] { "backupId", "createdAt", "sheet", "address", "formula", "formulaR1C1", "numberFormat", "locked", "spillAddress", "sourceRange" };
            for (var column = 1; column <= headers.Length; column++) backupApi.Cells.Item(2, column).Value2 = headers[column - 1];
            backupApi.Visible = 2;
            return backup;
        }
        finally { ComInterop.Release(sheets); }
    }

    private static void WriteBackup(dynamic sheet, int row, string backupId, dynamic cell, string sourceRange)
    {
        object? worksheet = null;
        try
        {
            worksheet = cell.Worksheet; dynamic worksheetApi = worksheet;
            var values = new object?[] { backupId, DateTimeOffset.UtcNow.ToString("O"), worksheetApi.Name, cell.Address[false, false], Formula(cell), FormulaR1C1(cell), cell.NumberFormat, Convert.ToBoolean(cell.Locked) ? "1" : "0", string.Empty, sourceRange };
            for (var column = 1; column <= values.Length; column++) sheet.Cells.Item(row, column).Value2 = values[column - 1];
        }
        finally { ComInterop.Release(worksheet); }
    }

    private static IEnumerable<BackupSummary> BackupSummaries(dynamic sheet)
    {
        var groups = new Dictionary<string, BackupSummary>(StringComparer.Ordinal);
        for (var row = 3; row <= LastBackupRow(sheet); row++)
        {
            string id = Convert.ToString((object?)sheet.Cells.Item(row, 1).Value2) ?? string.Empty;
            if (id.Length == 0) continue;
            if (!groups.TryGetValue(id, out var summary)) groups[id] = summary = new BackupSummary(id, Convert.ToString(sheet.Cells.Item(row, 2).Value2) ?? string.Empty);
            summary.FormulaCount++;
            summary.Sheets.Add(Convert.ToString(sheet.Cells.Item(row, 3).Value2) ?? string.Empty);
            summary.Ranges.Add(Convert.ToString(sheet.Cells.Item(row, 10).Value2) ?? string.Empty);
        }
        return groups.Values.OrderByDescending(summary => summary.CreatedAt);
    }

    private static int LastBackupRow(dynamic sheet)
    {
        object? cell = null;
        object? end = null;
        try { cell = sheet.Cells.Item(sheet.Rows.Count, 1); end = ((dynamic)cell).End(-4162); return Math.Max(2, Convert.ToInt32(((dynamic)end).Row)); }
        finally { ComInterop.Release(end); ComInterop.Release(cell); }
    }

    private static List<string[]> FindCycles(HashSet<string> nodes, List<FormulaEdge> edges)
    {
        var graph = edges.Where(edge => nodes.Contains(edge.To)).GroupBy(edge => edge.From, StringComparer.OrdinalIgnoreCase).ToDictionary(group => group.Key, group => group.Select(edge => edge.To).ToArray(), StringComparer.OrdinalIgnoreCase);
        var states = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase); var stack = new List<string>(); var cycles = new List<string[]>();
        void Visit(string node)
        {
            states[node] = 1; stack.Add(node);
            foreach (var next in graph.GetValueOrDefault(node, []))
            {
                if (!states.TryGetValue(next, out var state)) Visit(next);
                else if (state == 1) { var start = stack.FindIndex(item => item.Equals(next, StringComparison.OrdinalIgnoreCase)); if (start >= 0) cycles.Add([.. stack.Skip(start), next]); }
            }
            stack.RemoveAt(stack.Count - 1); states[node] = 2;
        }
        foreach (var node in nodes) if (!states.ContainsKey(node)) Visit(node);
        return cycles.DistinctBy(path => string.Join("->", path), StringComparer.OrdinalIgnoreCase).ToList();
    }

    private static string CellId(object cell)
    {
        object? worksheet = null;
        try { dynamic api = cell; worksheet = api.Worksheet; return $"{((dynamic)worksheet).Name}!{api.Address[false, false]}"; }
        finally { ComInterop.Release(worksheet); }
    }

    private static string Formula(dynamic cell) { try { return Convert.ToString(cell.Formula2) ?? string.Empty; } catch { return Convert.ToString(cell.Formula) ?? string.Empty; } }
    private static string FormulaR1C1(dynamic cell) { try { return Convert.ToString(cell.Formula2R1C1) ?? string.Empty; } catch { return Convert.ToString(cell.FormulaR1C1) ?? string.Empty; } }
    private static void SetFormula(object cell, string formula) =>
        ExcelFormulaWriter.Write(new ComExcelFormulaCell(cell), formula);
    private static string RemoveStringLiterals(string formula) => StringLiteralRegex().Replace(formula, "\"\"");
    private static string FormulaSummary(string operation) => operation switch { "repairFormulaReferences" => "已修复错误公式引用", "convertFormulasToValues" => "已备份公式并转换为值", "restoreFormulas" => "已恢复公式", "manageFormulaProtection" => "已更新公式保护", _ => "已检查公式依赖" };

    private sealed record FormulaNode(string Id, string Sheet, string Address, string Formula, object? Value);
    private sealed record FormulaEdge(string From, string To, string Kind, string Reference);
    private sealed record ProtectionInspection(List<object> Protection);
    private sealed class BackupSummary(string backupId, string createdAt)
    {
        public string BackupId { get; } = backupId; public string CreatedAt { get; } = createdAt; public int FormulaCount { get; set; }
        public HashSet<string> Sheets { get; } = []; public HashSet<string> Ranges { get; } = [];
    }
    private sealed class FormulaScope(object sheet, object range) : IDisposable
    {
        public object Sheet { get; } = sheet; public object Range { get; } = range;
        public void Dispose() { ComInterop.Release(Range); ComInterop.Release(Sheet); }
    }

    [GeneratedRegex("\"(?:[^\"]|\"\")*\"")]
    private static partial Regex StringLiteralRegex();
    [GeneratedRegex("\\[[^\\]]+\\](?:'(?<sheet>(?:[^']|'')+)'|(?<sheet>[^!]+))!(?<address>\\$?[A-Z]{1,3}\\$?\\d+(?::\\$?[A-Z]{1,3}\\$?\\d+)?)", RegexOptions.IgnoreCase)]
    private static partial Regex ExternalReferenceRegex();
    [GeneratedRegex("(?<![A-Za-z0-9_.])(?:'(?<quoted>(?:[^']|'')+)'|(?<plain>[A-Za-z_\\u4e00-\\u9fff][A-Za-z0-9_ .\\u4e00-\\u9fff-]*))!(?<address>\\$?[A-Z]{1,3}\\$?\\d+(?::\\$?[A-Z]{1,3}\\$?\\d+)?)", RegexOptions.IgnoreCase)]
    private static partial Regex QualifiedReferenceRegex();
    [GeneratedRegex("(?<![A-Za-z0-9_.!])(?<address>\\$?[A-Z]{1,3}\\$?\\d+(?::\\$?[A-Z]{1,3}\\$?\\d+)?)", RegexOptions.IgnoreCase)]
    private static partial Regex LocalReferenceRegex();
}
