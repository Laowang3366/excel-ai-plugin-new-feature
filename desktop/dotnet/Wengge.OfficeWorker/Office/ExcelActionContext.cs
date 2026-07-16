using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Office;

internal sealed class ExcelActionContext : IDisposable
{
    private readonly OfficeApplicationHandle? handle;
    private readonly OfficeDocumentLease? lease;
    private readonly bool borrowedWorkbook;
    private readonly bool openedWorkbook;

    public ExcelActionContext(OfficeApplicationProvider applications, OfficeActionRequest request)
    {
        var host = request.StringParam("host").ToLowerInvariant();
        var instanceId = request.StringParam("instanceId");
        if (instanceId.Length > 0)
        {
            lease = OfficeDocumentService.AcquireDocument("excel", request.FilePath, instanceId);
            OfficeHostRouting.Validate("excel", host, lease.Handle.ProgId);
            App = lease.Handle.Application;
            Workbook = lease.Handle.Document;
            borrowedWorkbook = true;
            return;
        }
        var progIds = OfficeHostRouting.ProgIds("excel", host);
        Trace("application.acquire.start");
        handle = string.IsNullOrWhiteSpace(request.FilePath)
            ? applications.GetActiveRequired(progIds, "当前没有已连接的 Excel/WPS 表格窗口")
            : applications.Create(progIds, "未找到可用的 Excel/WPS 表格 COM 应用");
        Trace($"application.acquire.done:{handle.ProgId}:created={handle.Created}");
        App = handle.Application;
        if (handle.Created)
        {
            try { App.Visible = false; } catch { }
            try { App.DisplayAlerts = false; App.EnableEvents = false; } catch { }
        }
        Trace("application.configure.done");
        if (string.IsNullOrWhiteSpace(request.FilePath))
        {
            Workbook = App.ActiveWorkbook ?? throw new OfficeWorkerException("document_not_found", "当前没有活动工作簿");
            return;
        }
        var wantedPath = Path.GetFullPath(request.FilePath);
        Trace("workbook.find.start");
        Workbook = FindWorkbook(wantedPath)!;
        Trace($"workbook.find.done:found={Workbook is not null}");
        if (Workbook is null)
        {
            object? workbooks = null;
            try
            {
                Trace("workbook.open.start");
                workbooks = App.Workbooks;
                Workbook = ((dynamic)workbooks).Open(wantedPath, 0, false, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, false, false, Type.Missing, false, true, Type.Missing);
                openedWorkbook = true;
                Trace("workbook.open.done");
            }
            finally { ComInterop.Release(workbooks); }
        }
    }

    public dynamic App { get; }
    public dynamic Workbook { get; private set; }
    public string ProgId => borrowedWorkbook ? lease!.Handle.ProgId : handle!.ProgId;

    public (object Sheet, object Range) GetRange(OfficeActionRequest request)
    {
        var (sheetName, address) = request.ExcelTarget();
        object sheet = Workbook.Worksheets.Item(sheetName);
        object range = ((dynamic)sheet).Range(address);
        return (sheet, range);
    }

    public void Save(OfficeActionRequest request)
    {
        if (!string.IsNullOrWhiteSpace(request.OutputPath) && !SamePath(request.OutputPath, request.FilePath))
        {
            var output = Path.GetFullPath(request.OutputPath);
            Directory.CreateDirectory(Path.GetDirectoryName(output) ?? Environment.CurrentDirectory);
            Workbook.SaveAs(output);
        }
        else
        {
            Workbook.Save();
        }
    }

    public void Dispose()
    {
        if (borrowedWorkbook)
        {
            Workbook = null!;
            lease?.Dispose();
            return;
        }
        if (Workbook is not null)
        {
            if (openedWorkbook)
            {
                try { Workbook.Close(false); } catch { }
            }
            ComInterop.Release(Workbook);
            Workbook = null!;
        }
        if (handle?.Created == true) { try { App.Quit(); } catch { } }
        handle?.Dispose();
        _ = handle?.WaitForExit();
    }

    private object? FindWorkbook(string path)
    {
        object? workbooks = null;
        try
        {
            workbooks = App.Workbooks;
            dynamic workbooksApi = workbooks;
            for (var index = 1; index <= Convert.ToInt32(workbooksApi.Count); index++)
            {
                object? candidate = workbooksApi.Item(index);
                try
                {
                    dynamic candidateApi = candidate!;
                    var candidatePath = Convert.ToString(candidateApi.FullName);
                    if (!string.IsNullOrWhiteSpace(candidatePath) && SamePath(candidatePath, path)) return candidate;
                }
                catch { }
                ComInterop.Release(candidate);
            }
            return null;
        }
        finally { ComInterop.Release(workbooks); }
    }

    private static bool SamePath(string? first, string? second) =>
        !string.IsNullOrWhiteSpace(first) && !string.IsNullOrWhiteSpace(second) &&
        string.Equals(Path.GetFullPath(first), Path.GetFullPath(second), StringComparison.OrdinalIgnoreCase);

    private static void Trace(string message)
    {
        if (Environment.GetEnvironmentVariable("WENGGE_OFFICE_SMOKE") == "1")
        {
            Console.Error.WriteLine($"[office-smoke] excel:{message}");
        }
    }
}
