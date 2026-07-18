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
        var activeProgIds = applications.DetectActiveProgIds([ExcelProgId, WpsProgId]);
        var activeHosts = activeProgIds.Select(HostForProgId).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        if (availableHosts.Count == 0 && activeHosts.Count == 0)
        {
            selectedHost = null;
            return new { connected = false, host = "unknown", availableHosts };
        }

        if (selectedHost is null && activeHosts.Count == 1)
        {
            selectedHost = activeHosts[0];
        }

        var active = selectedHost is null ? null : applications.TryGetActive(ProgIdsForHost(selectedHost));
        if (active is null)
        {
            var visibleHosts = activeHosts.Count > 0 ? activeHosts : availableHosts;
            return new
            {
                connected = false,
                host = visibleHosts.Count == 1 ? visibleHosts[0] : "unknown",
                availableHosts = visibleHosts,
                hostSelectionRequired = visibleHosts.Count > 1,
            };
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
                build = SafeString(() => app.Build),
                workbookName = SafeString(() => app.ActiveWorkbook?.Name),
                availableHosts,
                hostSelectionRequired = false,
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
        var progIds = ResolveProgIdsForActiveOperation();
        return applications.GetActiveRequired(progIds, "未连接到 Excel/WPS，请先打开应用并建立连接");
    }

    public OfficeApplicationHandle GetOrCreate()
    {
        var progIds = ResolveProgIdsForActiveOperation(allowNoActive: true);
        return applications.GetOrCreate(progIds, "无法启动 Excel/WPS");
    }

    private string[] ResolveProgIdsForActiveOperation(bool allowNoActive = false)
    {
        var activeProgIds = applications.DetectActiveProgIds([ExcelProgId, WpsProgId]);
        return ResolveProgIdsForActiveOperation(selectedHost, activeProgIds, allowNoActive);
    }

    internal static string[] ResolveProgIdsForActiveOperation(
        string? selectedHost,
        IReadOnlyCollection<string> activeProgIds,
        bool allowNoActive = false)
    {
        if (selectedHost is not null)
        {
            return ProgIdsForHost(selectedHost);
        }

        if (activeProgIds.Count > 1)
        {
            throw new OfficeWorkerException(
                "office_host_ambiguous",
                "同时检测到 Microsoft Excel 和 WPS 表格，请先选择目标宿主（excel.selectHost）。");
        }

        if (activeProgIds.Count == 1)
        {
            return [activeProgIds.First()];
        }

        if (allowNoActive)
        {
            return [ExcelProgId, WpsProgId];
        }

        return [ExcelProgId, WpsProgId];
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
