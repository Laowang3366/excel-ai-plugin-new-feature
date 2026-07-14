using System.Text.Json;
using Wengge.OfficeWorker.Office;
using Wengge.OfficeWorker.Protocol;
using Wengge.OfficeWorker.Runtime;

namespace Wengge.OfficeWorker.Tests;

public sealed class WorkerProtocolTests
{
    [Fact]
    public async Task HealthReportsCompatibleProtocol()
    {
        using var worker = OfficeWorkerHost.Create();
        var result = await worker.DispatchAsync(Request("worker.health"), CancellationToken.None);
        var json = JsonSerializer.SerializeToElement(result, JsonOptions);

        Assert.True(json.GetProperty("ready").GetBoolean());
        Assert.Equal(OfficeWorkerHost.ProtocolVersion, json.GetProperty("protocolVersion").GetInt32());
        Assert.Equal("x64", json.GetProperty("architecture").GetString());
    }

    [Fact]
    public async Task UnknownMethodReturnsStableError()
    {
        using var worker = OfficeWorkerHost.Create();
        var error = await Assert.ThrowsAsync<OfficeWorkerException>(() =>
            worker.DispatchAsync(Request("missing.method"), CancellationToken.None));

        Assert.Equal("method_not_found", error.Code);
    }

    [Fact]
    public async Task OfficeActionRequiresCompleteProtocolInput()
    {
        using var worker = OfficeWorkerHost.Create();
        var error = await Assert.ThrowsAsync<OfficeWorkerException>(() =>
            worker.DispatchAsync(Request("office.action.execute"), CancellationToken.None));

        Assert.Equal("invalid_params", error.Code);
    }

    [Fact]
    public async Task OfficeActionRejectsUnknownApplicationBeforeStartingCom()
    {
        using var worker = OfficeWorkerHost.Create();
        var error = await Assert.ThrowsAsync<OfficeWorkerException>(() =>
            worker.DispatchAsync(Request("office.action.execute", new
            {
                app = "outlook",
                action = "inspect",
                operation = "inspectFile",
            }), CancellationToken.None));

        Assert.Equal("unsupported_app", error.Code);
    }

    [Fact]
    public async Task OfficeDocumentRoutesValidateInputBeforeStartingCom()
    {
        using var worker = OfficeWorkerHost.Create();
        var error = await Assert.ThrowsAsync<OfficeWorkerException>(() =>
            worker.DispatchAsync(Request("office.documents.list", new { app = "outlook" }), CancellationToken.None));

        Assert.Equal("unsupported_app", error.Code);
    }

    [Fact]
    public async Task EmptyTransactionPreparationDoesNotInspectOfficeProcesses()
    {
        using var worker = OfficeWorkerHost.Create();
        var result = await worker.DispatchAsync(Request("office.transaction.prepare", new { filePaths = Array.Empty<string>() }), CancellationToken.None);
        var json = JsonSerializer.SerializeToElement(result, JsonOptions);

        Assert.Equal(JsonValueKind.Array, json.ValueKind);
        Assert.Empty(json.EnumerateArray());
    }

    [Fact]
    public async Task JsonRpcServerDispatchesIndependentRequestsConcurrentlyAndWritesCompleteResponses()
    {
        var first = JsonSerializer.Serialize(new { id = "first", method = "openxml.first", @params = new { } }, JsonOptions);
        var second = JsonSerializer.Serialize(new { id = "second", method = "openxml.second", @params = new { } }, JsonOptions);
        using var input = new StringReader(first + Environment.NewLine + second + Environment.NewLine);
        using var output = new StringWriter();
        using var error = new StringWriter();
        var bothStarted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var started = 0;
        var active = 0;
        var maximumActive = 0;
        var server = new JsonRpcServer(async (request, cancellationToken) =>
        {
            var current = Interlocked.Increment(ref active);
            UpdateMaximum(ref maximumActive, current);
            if (Interlocked.Increment(ref started) == 2) bothStarted.TrySetResult();
            await bothStarted.Task.WaitAsync(TimeSpan.FromSeconds(2), cancellationToken);
            await Task.Delay(25, cancellationToken);
            Interlocked.Decrement(ref active);
            return new { request.Id };
        }, input, output, error);

        await server.RunAsync(CancellationToken.None).WaitAsync(TimeSpan.FromSeconds(3));

        Assert.Equal(2, maximumActive);
        var responses = output.ToString().Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries)
            .Select(line => JsonDocument.Parse(line).RootElement.Clone()).ToArray();
        Assert.Equal(2, responses.Length);
        Assert.Equal(new[] { "first", "second" }, responses.Select(response => response.GetProperty("id").GetString()).Order().ToArray());
        Assert.All(responses, response => Assert.True(response.TryGetProperty("result", out _)));
    }

    [Fact]
    public void BuildsStablePerSectionCrossOfficeRequests()
    {
        var request = OfficeActionRequest.Parse(JsonSerializer.SerializeToElement(new
        {
            app = "excel",
            action = "insert",
            operation = "buildReportPackage",
            filePath = @"C:\reports\source.xlsx",
            @params = new
            {
                linked = true,
                overwrite = true,
                sections = new object[]
                {
                    new { linkId = "sales", sheetName = "Sales", range = "A1:B4", title = "销售" },
                    new { linkId = "cost", sheetName = "Cost", range = "C1:D4", title = "成本" },
                },
            },
        }, JsonOptions));

        var sections = ExcelCrossOfficeActionService.ReportSections(request);
        Assert.Equal(["sales", "cost"], sections.Select(section => section.LinkId));

        var sectionRequest = ExcelCrossOfficeActionService.ReportSectionRequest(
            request, sections[1], "exportRangeToWord", @"C:\reports\package.docx", "cost_word", updateExisting: true, firstSection: false);
        Assert.Equal("exportRangeToWord", sectionRequest.Operation);
        Assert.Equal("range:Cost!C1:D4", sectionRequest.Target);
        Assert.Equal("cost_word", sectionRequest.StringParam("linkId"));
        Assert.True(sectionRequest.BoolParam("updateExisting"));
        Assert.True(sectionRequest.BoolParam("allowMissingManaged"));
        Assert.False(sectionRequest.BoolParam("overwrite"));
        Assert.Equal("成本", sectionRequest.StringParam("title"));
    }

    [Fact]
    public void SourceRequestUsesSourceRoutingWithoutLeakingTargetRouting()
    {
        var request = OfficeActionRequest.Parse(JsonSerializer.SerializeToElement(new
        {
            app = "word",
            action = "insert",
            operation = "exportRangeToWord",
            filePath = @"C:\reports\target.docx",
            target = "bookmark:summary",
            @params = new
            {
                host = "kwps",
                instanceId = "word-instance",
                sourceHost = "ket",
                sourceInstanceId = "excel-instance",
                sheetName = "Sales",
            },
        }, JsonOptions));

        var source = OfficeHostRouting.SourceRequest(request, "excel", @"C:\reports\source.xlsx");

        Assert.Equal("excel", source.App);
        Assert.Equal(@"C:\reports\source.xlsx", source.FilePath);
        Assert.Null(source.Target);
        Assert.Equal("ket", source.StringParam("host"));
        Assert.Equal("excel-instance", source.StringParam("instanceId"));
        Assert.Equal("Sales", source.StringParam("sheetName"));
    }

    [Fact]
    public void DoneResultDoesNotClaimValidationWithoutRunningValidation()
    {
        var request = OfficeActionRequest.Parse(JsonSerializer.SerializeToElement(new
        {
            app = "excel",
            action = "edit",
            operation = "writeRange",
        }, JsonOptions));

        var result = JsonSerializer.SerializeToElement(OfficeActionResults.Done(request, "com", "done", null), JsonOptions);

        Assert.False(result.TryGetProperty("validation", out _));
    }

    [Fact]
    public void MissingPowerQueryMemberReportsUnsupportedHost()
    {
        var error = Assert.Throws<OfficeWorkerException>(() => ExcelQueryActionService.Snapshot(new MissingQueriesWorkbook(), string.Empty));

        Assert.Equal("power_query_unavailable", error.Code);
    }

    [Fact]
    public void PowerQueryReadFailureIsNotReportedAsUnsupportedHost()
    {
        var error = Assert.Throws<OfficeWorkerException>(() => ExcelQueryActionService.Snapshot(new FailingQueriesWorkbook(), string.Empty));

        Assert.Equal("power_query_inspection_failed", error.Code);
        Assert.IsType<InvalidOperationException>(error.InnerException);
    }

    private static void UpdateMaximum(ref int maximum, int candidate)
    {
        while (true)
        {
            var current = Volatile.Read(ref maximum);
            if (candidate <= current || Interlocked.CompareExchange(ref maximum, candidate, current) == current) return;
        }
    }

    private static RpcRequest Request(string method, object? parameters = null) =>
        new("test", method, JsonSerializer.SerializeToElement(parameters ?? new { }, JsonOptions));

    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);

    public sealed class MissingQueriesWorkbook;

    public sealed class FailingQueriesWorkbook
    {
        public object Queries => throw new InvalidOperationException("query collection is temporarily busy");
    }
}
