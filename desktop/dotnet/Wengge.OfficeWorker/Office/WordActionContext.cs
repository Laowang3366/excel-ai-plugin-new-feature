using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Office;

internal sealed class WordActionContext : IDisposable
{
    private readonly OfficeApplicationHandle? handle;
    private readonly OfficeDocumentLease? lease;
    private readonly WpsDocumentLaunch? wpsLaunch;
    private readonly bool borrowedDocument;
    private readonly bool openedDocument;
    private readonly string progId;

    public WordActionContext(OfficeApplicationProvider applications, OfficeActionRequest request)
    {
        var host = request.StringParam("host").ToLowerInvariant();
        var instanceId = request.StringParam("instanceId");
        if (instanceId.Length > 0)
        {
            lease = OfficeDocumentService.AcquireDocument("word", request.FilePath, instanceId);
            OfficeHostRouting.Validate("word", host, lease.Handle.ProgId);
            App = lease.Handle.Application;
            Document = lease.Handle.Document;
            progId = lease.Handle.ProgId;
            borrowedDocument = true;
            return;
        }
        if (!string.IsNullOrWhiteSpace(request.FilePath) && OfficeHostRouting.RequestsWps(host))
        {
            wpsLaunch = WpsFileLauncher.Open("word", request.FilePath);
            lease = wpsLaunch.Lease;
            App = lease.Handle.Application;
            Document = lease.Handle.Document;
            progId = lease.Handle.ProgId;
            borrowedDocument = true;
            return;
        }
        var progIds = OfficeHostRouting.ProgIds("word", host);
        handle = string.IsNullOrWhiteSpace(request.FilePath)
            ? applications.GetActiveRequired(progIds, "当前没有已连接的 Word/WPS 文字窗口")
            : applications.Create(progIds, "未找到可用的 Word/WPS 文字 COM 应用");
        progId = handle.ProgId;
        App = handle.Application;
        if (handle.Created) { try { App.Visible = false; App.DisplayAlerts = 0; } catch { } }
        if (string.IsNullOrWhiteSpace(request.FilePath))
        {
            Document = App.ActiveDocument ?? throw new OfficeWorkerException("document_not_found", "当前没有活动 Word 文档");
            return;
        }
        var wantedPath = Path.GetFullPath(request.FilePath);
        Document = FindDocument(wantedPath)!;
        if (Document is null)
        {
            object? documents = null;
            try
            {
                documents = App.Documents;
                Document = ((dynamic)documents).Open(wantedPath, false, false, false);
                openedDocument = true;
            }
            finally { ComInterop.Release(documents); }
        }
    }

    public WordActionContext(OfficeDocumentHandle documentHandle)
    {
        App = documentHandle.Application;
        Document = documentHandle.Document;
        progId = documentHandle.ProgId;
        borrowedDocument = true;
    }

    public dynamic App { get; }
    public dynamic Document { get; private set; }
    public string ProgId => progId;

    public void Save(OfficeActionRequest request)
    {
        if (!string.IsNullOrWhiteSpace(request.OutputPath) && !SamePath(request.OutputPath, request.FilePath))
        {
            var output = Path.GetFullPath(request.OutputPath);
            Directory.CreateDirectory(Path.GetDirectoryName(output) ?? Environment.CurrentDirectory);
            Document.SaveAs2(output);
        }
        else Document.Save();
    }

    public void Dispose()
    {
        if (borrowedDocument)
        {
            if (wpsLaunch is not null)
            {
                // Save() already persisted the requested mutation. Saving again during Close can make
                // Word normalize tracked-revision markup a second time and drop rebuilt bookmarks.
                try { Document.Close(0); } catch { }
                if (wpsLaunch.OwnsApplication) { try { App.Quit(); } catch { } }
            }
            Document = null!;
            if (wpsLaunch is not null) wpsLaunch.Dispose(); else lease?.Dispose();
            return;
        }
        if (Document is not null)
        {
            if (openedDocument) { try { Document.Close(0); } catch { } }
            ComInterop.Release(Document);
            Document = null!;
        }
        if (handle?.Created == true) { try { App.Quit(); } catch { } }
        handle?.Dispose();
        _ = handle?.WaitForExit();
    }

    private object? FindDocument(string path)
    {
        object? documents = null;
        try
        {
            documents = App.Documents; dynamic documentsApi = documents;
            for (var index = 1; index <= Convert.ToInt32(documentsApi.Count); index++)
            {
                object? candidate = documentsApi.Item(index);
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
        finally { ComInterop.Release(documents); }
    }

    private static bool SamePath(string? first, string? second) =>
        !string.IsNullOrWhiteSpace(first) && !string.IsNullOrWhiteSpace(second) && string.Equals(Path.GetFullPath(first), Path.GetFullPath(second), StringComparison.OrdinalIgnoreCase);
}
