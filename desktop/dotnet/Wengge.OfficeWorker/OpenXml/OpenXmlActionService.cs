using Wengge.OfficeWorker.Office;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.OpenXml;

internal sealed class OpenXmlActionService(
    OpenXmlExcelActionService excel,
    OpenXmlWordActionService word,
    OpenXmlPresentationActionService presentation)
{
    public object Execute(OfficeActionRequest request) => request.App switch
    {
        "excel" when OpenXmlExcelActionService.Supports(request.Operation) => excel.Execute(request),
        "word" when OpenXmlWordActionService.Supports(request.Operation) => word.Execute(request),
        "presentation" when OpenXmlPresentationActionService.Supports(request.Operation) => presentation.Execute(request),
        "excel" or "word" or "presentation" => OfficeActionResults.NeedsCom(request, $"Open XML 未覆盖 {request.App}/{request.Operation}，需要 COM 执行"),
        _ => throw new OfficeWorkerException("unsupported_app", $"不支持的 Office 应用: {request.App}"),
    };
}
