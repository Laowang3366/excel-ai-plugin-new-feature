using Wengge.OfficeWorker.Excel;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Tests;

public sealed class ExcelSessionRoutingTests
{
    [Fact]
    public void RequiresExplicitHostWhenBothComApplicationsAreActive()
    {
        var error = Assert.Throws<OfficeWorkerException>(() =>
            ExcelSessionService.ResolveProgIdsForActiveOperation(
                selectedHost: null,
                activeProgIds: ["Excel.Application", "Ket.Application"]));

        Assert.Equal("office_host_ambiguous", error.Code);
        Assert.Contains("excel.selectHost", error.Message);
    }

    [Fact]
    public void UsesExplicitHostEvenWhenTheOtherApplicationIsAlsoActive()
    {
        var result = ExcelSessionService.ResolveProgIdsForActiveOperation(
            selectedHost: "excel",
            activeProgIds: ["Excel.Application", "Ket.Application"]);

        Assert.Equal(["Excel.Application"], result);
    }

    [Fact]
    public void UsesTheOnlyActiveComApplicationWithoutPromptingForHost()
    {
        var result = ExcelSessionService.ResolveProgIdsForActiveOperation(
            selectedHost: null,
            activeProgIds: ["Ket.Application"]);

        Assert.Equal(["Ket.Application"], result);
    }
}
