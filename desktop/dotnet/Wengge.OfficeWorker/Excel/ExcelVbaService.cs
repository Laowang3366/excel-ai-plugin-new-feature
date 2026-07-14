using System.Reflection;
using System.Text.Json;
using System.Text.RegularExpressions;
using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Excel;

internal sealed partial class ExcelVbaService(ExcelSessionService sessions)
{
    public object DetectCapabilities()
    {
        try
        {
            using var handle = sessions.GetActiveRequired();
            dynamic app = handle.Application;
            object? workbook = null;
            object? project = null;
            object? components = null;
            try
            {
                workbook = app.ActiveWorkbook;
                if (workbook is null) throw new InvalidOperationException("当前没有活动工作簿");
                dynamic workbookApi = workbook;
                project = workbookApi.VBProject;
                dynamic projectApi = project;
                components = projectApi.VBComponents;
                _ = ((dynamic)components).Count;
                return new
                {
                    supported = true,
                    version = "VBA",
                    host = handle.ProgId == "Ket.Application" ? "wps" : "excel",
                };
            }
            finally
            {
                ComInterop.Release(components);
                ComInterop.Release(project);
                ComInterop.Release(workbook);
            }
        }
        catch (Exception exception)
        {
            return new
            {
                supported = false,
                reason = $"无法访问活动工作簿的 VBA 工程: {exception.Message}",
            };
        }
    }

    public object RunMacro(string macroName, JsonElement arguments)
    {
        if (string.IsNullOrWhiteSpace(macroName)) throw new OfficeWorkerException("invalid_params", "宏名称不能为空");
        var args = arguments.ValueKind == JsonValueKind.Array
            ? arguments.EnumerateArray().Select(ExcelValueConverter.FromJsonValue).ToArray()
            : [];
        if (args.Length > 30) throw new OfficeWorkerException("invalid_params", "宏参数不能超过 30 个");
        using var handle = sessions.GetActiveRequired();
        object app = handle.Application;
        var invokeArgs = new object?[args.Length + 1];
        invokeArgs[0] = macroName;
        Array.Copy(args, 0, invokeArgs, 1, args.Length);
        var result = app.GetType().InvokeMember(
            "Run",
            BindingFlags.InvokeMethod,
            binder: null,
            target: app,
            args: invokeArgs);
        return new { invoked = true, macroName, returnValue = result };
    }

    public object WriteModule(string moduleName, string code, string? entryPoint, bool save, string? saveAsPath)
    {
        var normalizedName = moduleName.Trim();
        var normalizedCode = NormalizeSource(code);
        if (!VbaIdentifierRegex().IsMatch(normalizedName) || normalizedName.Length > 31)
            throw new OfficeWorkerException("invalid_params", "模块名称必须是 1 到 31 位 VBA 标识符");
        if (normalizedCode.Length == 0) throw new OfficeWorkerException("invalid_params", "VBA 代码不能为空");
        if (!string.IsNullOrWhiteSpace(entryPoint) && !HasProcedure(normalizedCode, entryPoint))
            throw new OfficeWorkerException("invalid_params", $"VBA 代码中找不到入口过程: {entryPoint}");

        using var handle = sessions.GetActiveRequired();
        dynamic app = handle.Application;
        object? workbook = null;
        object? project = null;
        object? components = null;
        object? module = null;
        object? codeModule = null;
        var created = false;
        var oldCode = string.Empty;
        try
        {
            workbook = app.ActiveWorkbook ?? throw new InvalidOperationException("当前没有活动工作簿");
            dynamic workbookApi = workbook;
            project = workbookApi.VBProject;
            dynamic projectApi = project;
            components = projectApi.VBComponents;
            dynamic componentsApi = components;
            for (var index = 1; index <= Convert.ToInt32(componentsApi.Count); index++)
            {
                object? candidate = componentsApi.Item(index);
                dynamic candidateApi = candidate!;
                if (string.Equals(Convert.ToString(candidateApi.Name), normalizedName, StringComparison.OrdinalIgnoreCase))
                {
                    module = candidate;
                    break;
                }
                ComInterop.Release(candidate);
            }

            if (module is null)
            {
                module = componentsApi.Add(1);
                dynamic moduleApi = module;
                moduleApi.Name = normalizedName;
                created = true;
            }
            else if (Convert.ToInt32(((dynamic)module).Type) != 1)
            {
                throw new OfficeWorkerException("vba_component_conflict", $"同名组件不是标准模块: {normalizedName}");
            }

            codeModule = ((dynamic)module).CodeModule;
            dynamic codeModuleApi = codeModule;
            var oldLines = Convert.ToInt32(codeModuleApi.CountOfLines);
            if (oldLines > 0)
            {
                oldCode = Convert.ToString(codeModuleApi.Lines(1, oldLines)) ?? string.Empty;
                codeModuleApi.DeleteLines(1, oldLines);
            }
            codeModuleApi.AddFromString(normalizedCode);
            var lineCount = Convert.ToInt32(codeModuleApi.CountOfLines);
            var readBack = lineCount > 0 ? Convert.ToString(codeModuleApi.Lines(1, lineCount)) ?? string.Empty : string.Empty;
            if (!string.Equals(NormalizeSource(readBack), normalizedCode, StringComparison.Ordinal))
                throw new OfficeWorkerException("vba_verify_failed", "VBA 模块源码回读不一致");

            CompileProbe(app, workbookApi, componentsApi);
            if (save) SaveMacroWorkbook(app, workbookApi, saveAsPath);
            return new
            {
                moduleName = normalizedName,
                created,
                lineCount,
                sourceVerified = true,
                compileVerified = true,
                entryPoint,
                entryPointVerified = string.IsNullOrWhiteSpace(entryPoint) || HasProcedure(readBack, entryPoint),
                saved = save,
                workbookName = Convert.ToString(workbookApi.Name),
                workbookPath = Convert.ToString(workbookApi.FullName),
                host = handle.ProgId == "Ket.Application" ? "wps" : "excel",
            };
        }
        catch
        {
            try
            {
                if (module is not null && components is not null)
                {
                    dynamic componentsApi = components;
                    if (created) componentsApi.Remove(module);
                    else if (codeModule is not null)
                    {
                        dynamic codeModuleApi = codeModule;
                        var count = Convert.ToInt32(codeModuleApi.CountOfLines);
                        if (count > 0) codeModuleApi.DeleteLines(1, count);
                        if (oldCode.Length > 0) codeModuleApi.AddFromString(oldCode);
                    }
                }
            }
            catch { }
            throw;
        }
        finally
        {
            ComInterop.Release(codeModule);
            ComInterop.Release(module);
            ComInterop.Release(components);
            ComInterop.Release(project);
            ComInterop.Release(workbook);
        }
    }

    private static void CompileProbe(dynamic app, dynamic workbook, dynamic components)
    {
        object? probe = null;
        object? codeModule = null;
        try
        {
            probe = components.Add(1);
            dynamic probeApi = probe;
            var moduleName = $"WenggeProbe_{Guid.NewGuid():N}"[..24];
            probeApi.Name = moduleName;
            codeModule = probeApi.CodeModule;
            dynamic codeModuleApi = codeModule;
            codeModuleApi.AddFromString("Public Sub RunProbe()\r\nEnd Sub");
            var workbookName = (Convert.ToString(workbook.Name) ?? string.Empty).Replace("'", "''", StringComparison.Ordinal);
            app.Run($"'{workbookName}'!{moduleName}.RunProbe");
        }
        finally
        {
            if (probe is not null)
            {
                try { components.Remove(probe); } catch { }
            }
            ComInterop.Release(codeModule);
            ComInterop.Release(probe);
        }
    }

    private static void SaveMacroWorkbook(dynamic app, dynamic workbook, string? saveAsPath)
    {
        if (string.IsNullOrWhiteSpace(saveAsPath))
        {
            var extension = Path.GetExtension(Convert.ToString(workbook.Name) ?? string.Empty).ToLowerInvariant();
            if (extension is ".xlsm" or ".xlsb" or ".xls")
            {
                workbook.Save();
                return;
            }
            var directory = Convert.ToString(workbook.Path);
            if (string.IsNullOrWhiteSpace(directory))
                throw new OfficeWorkerException("macro_save_path_required", "请通过 saveAsPath 指定 .xlsm 保存路径");
            saveAsPath = Path.Combine(directory, $"{Path.GetFileNameWithoutExtension(Convert.ToString(workbook.Name))}-macro.xlsm");
        }
        var fullPath = Path.GetFullPath(saveAsPath);
        var extensionToSave = Path.GetExtension(fullPath).ToLowerInvariant();
        var format = extensionToSave switch
        {
            ".xlsm" => 52,
            ".xlsb" => 50,
            ".xls" => 56,
            _ => throw new OfficeWorkerException("invalid_macro_path", "宏工作簿必须保存为 .xlsm、.xlsb 或 .xls"),
        };
        var oldAlerts = app.DisplayAlerts;
        try
        {
            app.DisplayAlerts = false;
            workbook.SaveAs(fullPath, format);
        }
        finally
        {
            app.DisplayAlerts = oldAlerts;
        }
    }

    private static string NormalizeSource(string source) => source.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n').TrimStart().TrimEnd('\n');

    private static bool HasProcedure(string code, string entryPoint)
    {
        var name = entryPoint.Split('.').LastOrDefault()?.Trim();
        return !string.IsNullOrWhiteSpace(name) && Regex.IsMatch(code, $@"^\s*(?:Public\s+)?(?:Static\s+)?(?:Sub|Function)\s+{Regex.Escape(name)}\b", RegexOptions.IgnoreCase | RegexOptions.Multiline);
    }

    [GeneratedRegex("^[A-Za-z_][A-Za-z0-9_]*$")]
    private static partial Regex VbaIdentifierRegex();
}
