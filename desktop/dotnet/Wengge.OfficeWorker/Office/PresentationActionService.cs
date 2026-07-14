namespace Wengge.OfficeWorker.Office;

internal sealed class PresentationActionService(
    PresentationEditActionService edit,
    PresentationInspectionActionService inspect,
    PresentationBrandingActionService branding,
    PresentationLinkedContentActionService linked,
    PresentationPlaybackActionService playback)
{
    public object Execute(OfficeActionRequest request)
    {
        if (PresentationEditActionService.Supports(request.Operation)) return edit.Execute(request);
        if (PresentationInspectionActionService.Supports(request.Operation)) return inspect.Execute(request);
        if (PresentationBrandingActionService.Supports(request.Operation)) return branding.Execute(request);
        if (PresentationLinkedContentActionService.Supports(request.Operation)) return linked.Execute(request);
        return playback.Execute(request);
    }
}
