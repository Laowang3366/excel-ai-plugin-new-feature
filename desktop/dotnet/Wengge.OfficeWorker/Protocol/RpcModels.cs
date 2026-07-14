using System.Text.Json;
using System.Text.Json.Serialization;

namespace Wengge.OfficeWorker.Protocol;

public sealed record RpcRequest(
    [property: JsonPropertyName("id")] string Id,
    [property: JsonPropertyName("method")] string Method,
    [property: JsonPropertyName("params")] JsonElement Params);

public sealed record RpcError(
    [property: JsonPropertyName("code")] string Code,
    [property: JsonPropertyName("message")] string Message,
    [property: JsonPropertyName("data")] object? Data = null);

public sealed record RpcResponse(
    [property: JsonPropertyName("id")] string Id,
    [property: JsonPropertyName("result")] object? Result,
    [property: JsonPropertyName("error")] RpcError? Error)
{
    public static RpcResponse Success(string id, object? result) => new(id, result, null);

    public static RpcResponse Failure(string id, string code, string message, object? data = null) =>
        new(id, null, new RpcError(code, message, data));
}

public sealed class OfficeWorkerException : Exception
{
    public OfficeWorkerException(string code, string message, object? details = null, Exception? inner = null)
        : base(message, inner)
    {
        Code = code;
        Details = details;
    }

    public string Code { get; }

    public object? Details { get; }
}
