using Wengge.OfficeWorker.Office;

namespace Wengge.OfficeWorker.Tests;

public sealed class OfficeDocumentStatusTests
{
    [Fact]
    public void ApplicationMetadataEnrichesRotDocumentWithWpsProcess()
    {
        var document = new FakeDocument { Name = "status.pptx" };
        var application = new FakeApplication { Version = "12.0" };
        var rotHandle = new OfficeDocumentHandle(
            "presentation", document, application, "presentation:rot:test", 0, 0,
            "microsoft-office", "PowerPoint.Application");
        var applicationHandle = new OfficeDocumentHandle(
            "presentation", document, application, "presentation:50864:100", 50864, 100,
            "wps", "Wpp.Application");

        var merged = OfficeDocumentService.MergeDuplicateMetadata(rotHandle, applicationHandle);

        Assert.Equal(50864, merged.ProcessId);
        Assert.Equal(100, merged.Hwnd);
        Assert.Equal("wps", merged.Host);
        Assert.Equal("Wpp.Application", merged.ProgId);
        Assert.Equal("presentation:50864:100", merged.InstanceId);
    }

    [Fact]
    public void PresentationStatusNormalizesWppHostAndIncludesProcess()
    {
        var handle = new OfficeDocumentHandle(
            "presentation",
            new FakeDocument { Name = "status.pptx" },
            new FakeApplication { Version = "12.0" },
            "presentation:50864:100",
            50864,
            100,
            "wps",
            "Wpp.Application");

        var status = OfficeDocumentService.BuildConnectionStatus("presentation", handle);

        Assert.True(status.Connected);
        Assert.Equal("wps", status.Host);
        Assert.Equal("12.0", status.Version);
        Assert.Equal("status.pptx", status.PresentationName);
        Assert.Equal(50864, status.ProcessId);
    }

    public sealed class FakeDocument
    {
        public string Name { get; init; } = string.Empty;
    }

    public sealed class FakeApplication
    {
        public string Version { get; init; } = string.Empty;
    }
}
