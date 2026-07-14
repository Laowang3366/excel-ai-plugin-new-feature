using System.Diagnostics;
using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Excel;

internal sealed class ExcelSessionService(OfficeApplicationProvider applications)
{
    private const string ExcelProgId = "Excel.Application";
    private const string WpsProgId = "Ket.Application";
    private string? selectedHost;

    public object DetectStatus()
    {
        var availableHosts = DetectProcesses();
        if (availableHosts.Count == 0)
        {
            selectedHost = null;
            return new { connected = false, host = "unknown", availableHosts };
        }

        var active = TryGetActive(availableHosts);
        if (active is null)
        {
            return new { connected = false, host = availableHosts.Count == 1 ? availableHosts[0] : "unknown", availableHosts };
        }

        using (active)
        {
            dynamic app = active.Application;
            var host = HostForProgId(active.ProgId);
            selectedHost = host;
            return new
            {
                connected = true,
                host,
                version = SafeString(() => app.Version),
                workbookName = SafeString(() => app.ActiveWorkbook?.Name),
                availableHosts,
            };
        }
    }

    public object Connect() => DetectStatus();

    public object SelectHost(string host)
    {
        if (host is not ("excel" or "wps"))
        {
            throw new OfficeWorkerException("invalid_params", "host 只能是 excel 或 wps");
        }

        selectedHost = host;
        using var active = applications.TryGetActive(ProgIdsForHost(host));
        if (active is null)
        {
            return new { connected = false, host };
        }

        dynamic app = active.Application;
        return new
        {
            connected = true,
            host,
            version = SafeString(() => app.Version),
            workbookName = SafeString(() => app.ActiveWorkbook?.Name),
        };
    }

    public OfficeApplicationHandle GetActiveRequired()
    {
        var progIds = selectedHost is null
            ? new[] { ExcelProgId, WpsProgId }
            : ProgIdsForHost(selectedHost);
        return applications.GetActiveRequired(progIds, "未连接到 Excel/WPS，请先打开应用并建立连接");
    }

    public OfficeApplicationHandle GetOrCreate()
    {
        var progIds = selectedHost is null
            ? new[] { ExcelProgId, WpsProgId }
            : ProgIdsForHost(selectedHost);
        return applications.GetOrCreate(progIds, "无法启动 Excel/WPS");
    }

    private OfficeApplicationHandle? TryGetActive(IReadOnlyCollection<string> availableHosts)
    {
        if (selectedHost is not null && availableHosts.Contains(selectedHost))
        {
            var selected = applications.TryGetActive(ProgIdsForHost(selectedHost));
            if (selected is not null)
            {
                return selected;
            }
        }

        return availableHosts.Count == 1
            ? applications.TryGetActive(ProgIdsForHost(availableHosts.First()))
            : null;
    }

    private static List<string> DetectProcesses()
    {
        var hosts = new List<string>();
        if (Process.GetProcessesByName("EXCEL").Length > 0)
        {
            hosts.Add("excel");
        }

        if (Process.GetProcessesByName("et").Length > 0 || Process.GetProcessesByName("wps").Length > 0)
        {
            hosts.Add("wps");
        }

        return hosts;
    }

    private static string[] ProgIdsForHost(string host) => host == "wps" ? [WpsProgId] : [ExcelProgId];

    private static string HostForProgId(string progId) => progId == WpsProgId ? "wps" : "excel";

    private static string? SafeString(Func<object?> value)
    {
        try
        {
            return Convert.ToString(value(), System.Globalization.CultureInfo.InvariantCulture);
        }
        catch
        {
            return null;
        }
    }
}
