using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Office;

internal sealed class PresentationActionContext : IDisposable
{
    private readonly OfficeApplicationHandle? handle;
    private readonly OfficeDocumentLease? lease;
    private readonly WpsDocumentLaunch? wpsLaunch;
    private readonly bool borrowedPresentation;
    private readonly bool openedPresentation;

    public PresentationActionContext(OfficeApplicationProvider applications, OfficeActionRequest request)
    {
        var host = request.StringParam("host").ToLowerInvariant();
        var instanceId = request.StringParam("instanceId");
        if (instanceId.Length > 0)
        {
            lease = OfficeDocumentService.AcquireDocument("presentation", request.FilePath, instanceId);
            OfficeHostRouting.Validate("presentation", host, lease.Handle.ProgId);
            App = lease.Handle.Application;
            Presentation = lease.Handle.Document;
            borrowedPresentation = true;
            return;
        }
        if (!string.IsNullOrWhiteSpace(request.FilePath) && OfficeHostRouting.RequestsWps(host))
        {
            wpsLaunch = WpsFileLauncher.Open("presentation", request.FilePath);
            lease = wpsLaunch.Lease;
            App = lease.Handle.Application;
            Presentation = lease.Handle.Document;
            borrowedPresentation = true;
            return;
        }
        var progIds = OfficeHostRouting.ProgIds("presentation", host);
        handle = string.IsNullOrWhiteSpace(request.FilePath)
            ? applications.GetActiveRequired(progIds, "当前没有已连接的 PowerPoint/WPS 演示窗口")
            : applications.Create(progIds, "未找到可用的 PowerPoint/WPS 演示 COM 应用");
        App = handle.Application;
        if (handle.Created) { try { App.Visible = -1; App.DisplayAlerts = 1; } catch { } }
        if (string.IsNullOrWhiteSpace(request.FilePath))
        {
            Presentation = App.ActivePresentation ?? throw new OfficeWorkerException("document_not_found", "当前没有活动演示文稿");
            return;
        }
        var wantedPath = Path.GetFullPath(request.FilePath);
        Presentation = FindPresentation(wantedPath)!;
        if (Presentation is null)
        {
            object? presentations = null;
            try { presentations = App.Presentations; Presentation = ((dynamic)presentations).Open(wantedPath); openedPresentation = true; }
            finally { ComInterop.Release(presentations); }
        }
    }

    public dynamic App { get; }
    public dynamic Presentation { get; private set; }
    public string ProgId => borrowedPresentation ? lease!.Handle.ProgId : handle!.ProgId;

    public object GetSlide(int index)
    {
        object? slides = null;
        try
        {
            slides = Presentation.Slides; dynamic slidesApi = slides;
            if (index < 1 || index > Convert.ToInt32(slidesApi.Count)) throw new OfficeWorkerException("slide_not_found", $"幻灯片序号超出范围: {index}");
            return slidesApi.Item(index);
        }
        finally { ComInterop.Release(slides); }
    }

    public void Save(OfficeActionRequest request)
    {
        if (!string.IsNullOrWhiteSpace(request.OutputPath) && !SamePath(request.OutputPath, request.FilePath))
        {
            var output = Path.GetFullPath(request.OutputPath); Directory.CreateDirectory(Path.GetDirectoryName(output) ?? Environment.CurrentDirectory); Presentation.SaveAs(output);
        }
        else Presentation.Save();
    }

    public void Dispose()
    {
        if (borrowedPresentation)
        {
            if (wpsLaunch is not null)
            {
                try { Presentation.Close(); } catch { }
                if (wpsLaunch.OwnsApplication) { try { App.Quit(); } catch { } }
            }
            Presentation = null!;
            if (wpsLaunch is not null) wpsLaunch.Dispose(); else lease?.Dispose();
            return;
        }
        if (Presentation is not null)
        {
            if (openedPresentation) { try { Presentation.Close(); } catch { } }
            ComInterop.Release(Presentation); Presentation = null!;
        }
        if (handle?.Created == true) { try { App.Quit(); } catch { } }
        handle?.Dispose();
        _ = handle?.WaitForExit();
    }

    private object? FindPresentation(string path)
    {
        object? presentations = null;
        try
        {
            presentations = App.Presentations; dynamic collection = presentations;
            for (var index = 1; index <= Convert.ToInt32(collection.Count); index++)
            {
                object? candidate = collection.Item(index);
                try
                {
                    var candidatePath = Convert.ToString(((dynamic)candidate).FullName);
                    if (!string.IsNullOrWhiteSpace(candidatePath) && SamePath(candidatePath, path)) return candidate;
                }
                catch { }
                ComInterop.Release(candidate);
            }
            return null;
        }
        finally { ComInterop.Release(presentations); }
    }

    private static bool SamePath(string? first, string? second) =>
        !string.IsNullOrWhiteSpace(first) && !string.IsNullOrWhiteSpace(second) && string.Equals(Path.GetFullPath(first), Path.GetFullPath(second), StringComparison.OrdinalIgnoreCase);
}
