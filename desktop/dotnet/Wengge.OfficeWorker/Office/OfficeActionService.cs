using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Office;

internal sealed class OfficeActionService(
    ExcelActionService excel,
    WordActionService word,
    PresentationActionService presentation)
{
    public object Execute(OfficeActionRequest request) => request.App switch
    {
        "excel" => excel.Execute(request),
        "word" => word.Execute(request),
        "presentation" => presentation.Execute(request),
        _ => throw new OfficeWorkerException("unsupported_app", $"不支持的 Office 应用: {request.App}"),
    };
}
