using Wengge.OfficeWorker.Com;

namespace Wengge.OfficeWorker.Office;

internal sealed class ExcelWorkbookActionService(
    ExcelChartActionService charts,
    ExcelObjectActionService objects,
    ExcelTemplatePrintActionService templates)
{
    private static readonly HashSet<string> Operations =
    [
        "inspectCharts", "formatChart",
        "inspectWorkbookObjects", "manageWorkbookObject", "manageWorksheetObjects",
        "captureWorkbookTemplate", "inspectWorkbookFormatting", "applyWorkbookTemplate",
        "inspectPrintSettings", "configurePrint",
    ];

    public static bool Supports(string operation) => Operations.Contains(operation);

    public object Execute(OfficeActionRequest request)
    {
        if (ExcelChartActionService.Supports(request.Operation)) return charts.Execute(request);
        if (ExcelObjectActionService.Supports(request.Operation)) return objects.Execute(request);
        return templates.Execute(request);
    }
}
