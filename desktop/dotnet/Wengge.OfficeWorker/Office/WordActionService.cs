using Wengge.OfficeWorker.Com;

namespace Wengge.OfficeWorker.Office;

internal sealed class WordActionService(
    WordFormattingActionService formatting,
    WordReferenceRevisionActionService references,
    WordMailMergeContentActionService content,
    WordLinkedContentActionService linked)
{
    public object Execute(OfficeActionRequest request)
    {
        if (WordFormattingActionService.Supports(request.Operation)) return formatting.Execute(request);
        if (WordReferenceRevisionActionService.Supports(request.Operation)) return references.Execute(request);
        if (WordMailMergeContentActionService.Supports(request.Operation)) return content.Execute(request);
        return linked.Execute(request);
    }
}
