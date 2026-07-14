using System.Text.Json;
using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Excel;

internal sealed class ExcelWorkbookService(ExcelSessionService sessions)
{
    public object Inspect()
    {
        using var handle = sessions.GetActiveRequired();
        dynamic app = handle.Application;
        object? workbook = null;
        object? sheets = null;
        try
        {
            workbook = app.ActiveWorkbook ?? throw new InvalidOperationException("当前没有活动工作簿");
            dynamic workbookApi = workbook;
            sheets = workbookApi.Sheets;
            dynamic sheetsApi = sheets;
            var sheetItems = new List<object>();
            for (var index = 1; index <= Convert.ToInt32(sheetsApi.Count); index++)
            {
                object? sheet = null;
                object? usedRange = null;
                try
                {
                    sheet = sheetsApi.Item(index);
                    dynamic sheetApi = sheet;
                    usedRange = sheetApi.UsedRange;
                    dynamic usedRangeApi = usedRange;
                    sheetItems.Add(new
                    {
                        name = Convert.ToString(sheetApi.Name),
                        index,
                        rows = Convert.ToInt32(usedRangeApi.Rows.Count),
                        columns = Convert.ToInt32(usedRangeApi.Columns.Count),
                        visible = Convert.ToInt32(sheetApi.Visible) != 0,
                    });
                }
                finally
                {
                    ComInterop.Release(usedRange);
                    ComInterop.Release(sheet);
                }
            }

            return new
            {
                name = Convert.ToString(workbookApi.Name),
                path = Convert.ToString(workbookApi.FullName),
                saved = Convert.ToBoolean(workbookApi.Saved),
                host = handle.ProgId == "Ket.Application" ? "wps" : "excel",
                sheets = sheetItems,
            };
        }
        finally
        {
            ComInterop.Release(sheets);
            ComInterop.Release(workbook);
        }
    }

    public object Open(string filePath)
    {
        var fullPath = Path.GetFullPath(filePath);
        if (!File.Exists(fullPath))
        {
            throw new OfficeWorkerException("file_not_found", $"工作簿不存在: {fullPath}");
        }

        using var handle = sessions.GetOrCreate();
        dynamic app = handle.Application;
        object? workbooks = null;
        object? workbook = null;
        try
        {
            app.Visible = true;
            workbooks = app.Workbooks;
            dynamic workbooksApi = workbooks;
            workbook = workbooksApi.Open(fullPath);
            dynamic workbookApi = workbook;
            return new { success = true, workbookName = Convert.ToString(workbookApi.Name) };
        }
        finally
        {
            ComInterop.Release(workbook);
            ComInterop.Release(workbooks);
        }
    }

    public object Create(string filePath, JsonElement sheetNames)
    {
        var fullPath = Path.GetFullPath(filePath);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath) ?? Environment.CurrentDirectory);
        using var handle = sessions.GetOrCreate();
        dynamic app = handle.Application;
        object? workbooks = null;
        object? workbook = null;
        object? sheets = null;
        try
        {
            app.Visible = true;
            workbooks = app.Workbooks;
            dynamic workbooksApi = workbooks;
            workbook = workbooksApi.Add();
            dynamic workbookApi = workbook;
            sheets = workbookApi.Sheets;
            dynamic sheetsApi = sheets;
            var names = sheetNames.ValueKind == JsonValueKind.Array
                ? sheetNames.EnumerateArray().Select(item => item.GetString()).Where(name => !string.IsNullOrWhiteSpace(name)).ToArray()
                : [];
            for (var index = 0; index < names.Length; index++)
            {
                object? sheet = null;
                try
                {
                    if (index < Convert.ToInt32(sheetsApi.Count))
                    {
                        sheet = sheetsApi.Item(index + 1);
                    }
                    else
                    {
                        sheet = sheetsApi.Add(After: sheetsApi.Item(sheetsApi.Count));
                    }

                    dynamic sheetApi = sheet;
                    sheetApi.Name = names[index];
                }
                finally
                {
                    ComInterop.Release(sheet);
                }
            }

            workbookApi.SaveAs(fullPath);
            return new { success = true, workbookName = Convert.ToString(workbookApi.Name) };
        }
        finally
        {
            ComInterop.Release(sheets);
            ComInterop.Release(workbook);
            ComInterop.Release(workbooks);
        }
    }

    public object Save(string? saveAsPath)
    {
        using var handle = sessions.GetActiveRequired();
        dynamic app = handle.Application;
        object? workbook = null;
        try
        {
            workbook = app.ActiveWorkbook ?? throw new InvalidOperationException("当前没有活动工作簿");
            dynamic workbookApi = workbook;
            if (string.IsNullOrWhiteSpace(saveAsPath))
            {
                workbookApi.Save();
            }
            else
            {
                var fullPath = Path.GetFullPath(saveAsPath);
                Directory.CreateDirectory(Path.GetDirectoryName(fullPath) ?? Environment.CurrentDirectory);
                workbookApi.SaveAs(fullPath);
            }

            return new { success = true };
        }
        finally
        {
            ComInterop.Release(workbook);
        }
    }

    public object Switch(string workbookName)
    {
        using var handle = sessions.GetActiveRequired();
        dynamic app = handle.Application;
        object? workbooks = null;
        object? workbook = null;
        try
        {
            workbooks = app.Workbooks;
            dynamic workbooksApi = workbooks;
            workbook = workbooksApi.Item(workbookName);
            dynamic workbookApi = workbook;
            workbookApi.Activate();
            return new { success = true, workbookName = Convert.ToString(workbookApi.Name) };
        }
        finally
        {
            ComInterop.Release(workbook);
            ComInterop.Release(workbooks);
        }
    }

    public object SheetOperation(string operation, string sheetName, JsonElement options)
    {
        using var handle = sessions.GetActiveRequired();
        dynamic app = handle.Application;
        object? workbook = null;
        object? sheets = null;
        object? sheet = null;
        try
        {
            workbook = app.ActiveWorkbook ?? throw new InvalidOperationException("当前没有活动工作簿");
            dynamic workbookApi = workbook;
            sheets = workbookApi.Sheets;
            dynamic sheetsApi = sheets;
            if (operation == "add")
            {
                sheet = sheetsApi.Add();
                dynamic added = sheet;
                added.Name = sheetName;
                return new { operation, sheetName };
            }

            sheet = sheetsApi.Item(sheetName);
            dynamic sheetApi = sheet;
            switch (operation)
            {
                case "rename":
                    sheetApi.Name = options.GetProperty("newName").GetString();
                    break;
                case "delete":
                    app.DisplayAlerts = false;
                    sheetApi.Delete();
                    app.DisplayAlerts = true;
                    break;
                case "copy":
                    sheetApi.Copy(After: sheetsApi.Item(sheetsApi.Count));
                    break;
                case "move":
                    var position = options.TryGetProperty("position", out var rawPosition) ? rawPosition.GetInt32() : 1;
                    sheetApi.Move(Before: sheetsApi.Item(Math.Max(1, position)));
                    break;
                default:
                    throw new OfficeWorkerException("unsupported_operation", $"不支持工作表操作: {operation}");
            }

            return new { operation, sheetName };
        }
        finally
        {
            ComInterop.Release(sheet);
            ComInterop.Release(sheets);
            ComInterop.Release(workbook);
        }
    }
}
