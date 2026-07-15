using Wengge.OfficeWorker.Office;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Tests;

public sealed class WordFormattingPatternTests
{
    [Fact]
    public void MatchesHeadingPatternAcceptsValidExpressions()
    {
        Assert.True(WordFormattingActionService.MatchesHeadingPattern("1.2 Revenue", @"^\d+\.\d+"));
        Assert.False(WordFormattingActionService.MatchesHeadingPattern("Revenue", @"^\d+\.\d+"));
        Assert.False(WordFormattingActionService.MatchesHeadingPattern("Revenue", string.Empty));
    }

    [Fact]
    public void MatchesHeadingPatternRejectsInvalidExpressions()
    {
        var error = Assert.Throws<OfficeWorkerException>(() =>
            WordFormattingActionService.MatchesHeadingPattern("Revenue", "["));

        Assert.Equal("invalid_params", error.Code);
        Assert.Contains("无效", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void MatchesHeadingPatternBoundsCatastrophicBacktracking()
    {
        var text = new string('a', 50_000) + "!";
        var error = Assert.Throws<OfficeWorkerException>(() =>
            WordFormattingActionService.MatchesHeadingPattern(text, "^(a+)+$"));

        Assert.Equal("invalid_params", error.Code);
        Assert.Contains("超时", error.Message, StringComparison.Ordinal);
    }
}
