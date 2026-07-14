using System.Text.Json;

namespace Wengge.OfficeWorker.Protocol;

public sealed class JsonRpcServer
{
    private static readonly JsonSerializerOptions SerializerOptions = new(JsonSerializerDefaults.Web)
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull,
    };

    private readonly Func<RpcRequest, CancellationToken, Task<object?>> dispatch;
    private readonly TextReader input;
    private readonly TextWriter output;
    private readonly TextWriter error;
    private readonly SemaphoreSlim outputGate = new(1, 1);

    public JsonRpcServer(
        Func<RpcRequest, CancellationToken, Task<object?>> dispatch,
        TextReader input,
        TextWriter output,
        TextWriter error)
    {
        this.dispatch = dispatch;
        this.input = input;
        this.output = output;
        this.error = error;
    }

    public async Task RunAsync(CancellationToken cancellationToken)
    {
        var inFlight = new List<Task>();
        while (!cancellationToken.IsCancellationRequested)
        {
            var line = await input.ReadLineAsync(cancellationToken);
            if (line is null)
            {
                break;
            }

            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            inFlight.Add(HandleAndWriteAsync(line, cancellationToken));
            var completed = inFlight.Where(task => task.IsCompleted).ToArray();
            if (completed.Length == 0) continue;
            await Task.WhenAll(completed);
            foreach (var task in completed) inFlight.Remove(task);
        }
        await Task.WhenAll(inFlight);
    }

    private async Task HandleAndWriteAsync(string line, CancellationToken cancellationToken)
    {
        var response = await HandleLineAsync(line, cancellationToken);
        var serialized = JsonSerializer.Serialize(response, SerializerOptions);
        await outputGate.WaitAsync(cancellationToken);
        try
        {
            await output.WriteLineAsync(serialized);
            await output.FlushAsync(cancellationToken);
        }
        finally { outputGate.Release(); }
    }

    internal async Task<RpcResponse> HandleLineAsync(string line, CancellationToken cancellationToken)
    {
        RpcRequest? request = null;
        try
        {
            request = JsonSerializer.Deserialize<RpcRequest>(line, SerializerOptions);
            if (request is null || string.IsNullOrWhiteSpace(request.Id) || string.IsNullOrWhiteSpace(request.Method))
            {
                return RpcResponse.Failure(request?.Id ?? string.Empty, "invalid_request", "RPC 请求缺少 id 或 method");
            }

            var result = await dispatch(request, cancellationToken);
            return RpcResponse.Success(request.Id, result);
        }
        catch (JsonException exception)
        {
            return RpcResponse.Failure(request?.Id ?? string.Empty, "invalid_json", exception.Message);
        }
        catch (OfficeWorkerException exception)
        {
            return RpcResponse.Failure(request?.Id ?? string.Empty, exception.Code, exception.Message, exception.Details);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            return RpcResponse.Failure(request?.Id ?? string.Empty, "cancelled", "操作已取消");
        }
        catch (Exception exception)
        {
            await error.WriteLineAsync($"[{DateTimeOffset.Now:O}] {exception}");
            await error.FlushAsync(cancellationToken);
            return RpcResponse.Failure(request?.Id ?? string.Empty, "worker_error", exception.Message);
        }
    }
}
