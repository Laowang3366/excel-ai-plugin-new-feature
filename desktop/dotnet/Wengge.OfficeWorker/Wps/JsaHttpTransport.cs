using System.Net;
using System.Text;
using System.Text.Json;
using Wengge.OfficeWorker.Protocol;

namespace Wengge.OfficeWorker.Wps;

internal sealed class JsaHttpTransport : IDisposable
{
    private readonly object sync = new();
    private HttpListener? listener;
    private CancellationTokenSource? listenerCancellation;
    private PendingCommand? pending;
    private string token = string.Empty;
    private DateTimeOffset lastPoll = DateTimeOffset.MinValue;

    public Task StartAsync(string bridgeToken, CancellationToken cancellationToken)
    {
        lock (sync)
        {
            token = bridgeToken;
            if (listener is not null) return Task.CompletedTask;
            listener = new HttpListener();
            listener.Prefixes.Add("http://127.0.0.1:45221/");
            listener.Start();
            listenerCancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            _ = Task.Run(() => ListenAsync(listener, listenerCancellation.Token), CancellationToken.None);
            return Task.CompletedTask;
        }
    }

    public async Task<bool> WaitForClientAsync(int timeoutMs, CancellationToken cancellationToken)
    {
        var deadline = DateTimeOffset.UtcNow.AddMilliseconds(timeoutMs);
        while (DateTimeOffset.UtcNow < deadline)
        {
            if (DateTimeOffset.UtcNow - lastPoll < TimeSpan.FromSeconds(2)) return true;
            await Task.Delay(100, cancellationToken);
        }
        return false;
    }

    public async Task<JsonElement> SendAsync(string type, object? data, CancellationToken cancellationToken)
    {
        if (listener is null) throw new OfficeWorkerException("jsa_transport_stopped", "WPS JSA 本地桥接服务尚未启动");
        if (!await WaitForClientAsync(2_000, cancellationToken))
            throw new OfficeWorkerException("jsa_not_connected", "WPS JSA 加载项未连接，请完全退出并重新打开 WPS 表格");
        var command = new PendingCommand(Guid.NewGuid().ToString("N"), type, data);
        lock (sync)
        {
            if (pending is not null) throw new OfficeWorkerException("jsa_busy", "已有 WPS JSA 写入任务正在执行");
            pending = command;
        }
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(TimeSpan.FromSeconds(15));
        try
        {
            return await command.Completion.Task.WaitAsync(timeout.Token);
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            throw new OfficeWorkerException("jsa_timeout", "WPS JSA 内部宏操作超时");
        }
        finally
        {
            lock (sync) { if (ReferenceEquals(pending, command)) pending = null; }
        }
    }

    public void Dispose()
    {
        listenerCancellation?.Cancel();
        listener?.Close();
        listenerCancellation?.Dispose();
        listenerCancellation = null;
        listener = null;
    }

    private async Task ListenAsync(HttpListener activeListener, CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            try
            {
                var context = await activeListener.GetContextAsync().WaitAsync(cancellationToken);
                _ = Task.Run(() => HandleAsync(context, cancellationToken), CancellationToken.None);
            }
            catch (OperationCanceledException) { return; }
            catch (HttpListenerException) when (cancellationToken.IsCancellationRequested) { return; }
        }
    }

    private async Task HandleAsync(HttpListenerContext context, CancellationToken cancellationToken)
    {
        var response = context.Response;
        response.Headers["Access-Control-Allow-Origin"] = "*";
        response.Headers["Access-Control-Allow-Headers"] = "Content-Type, X-Wengge-Token";
        if (context.Request.HttpMethod == "OPTIONS") { response.StatusCode = 204; response.Close(); return; }
        if (context.Request.Headers["X-Wengge-Token"] != token) { response.StatusCode = 403; response.Close(); return; }
        if (context.Request.HttpMethod == "GET" && context.Request.Url?.AbsolutePath == "/command")
        {
            lastPoll = DateTimeOffset.UtcNow;
            PendingCommand? command;
            lock (sync)
            {
                command = pending is { Delivered: false } ? pending : null;
                if (command is not null) command.Delivered = true;
            }
            if (command is null) { response.StatusCode = 204; response.Close(); return; }
            await WriteJsonAsync(response, new { id = command.Id, type = command.Type, data = command.Data }, cancellationToken);
            return;
        }
        if (context.Request.HttpMethod == "POST" && context.Request.Url?.AbsolutePath == "/response")
        {
            var body = await JsonSerializer.DeserializeAsync<JsonElement>(context.Request.InputStream, cancellationToken: cancellationToken);
            PendingCommand? command;
            lock (sync) { command = pending; }
            if (command is null || !body.TryGetProperty("id", out var id) || id.GetString() != command.Id)
            {
                response.StatusCode = 409; response.Close(); return;
            }
            if (body.TryGetProperty("ok", out var ok) && ok.GetBoolean())
                command.Completion.TrySetResult(body.TryGetProperty("result", out var result) ? result.Clone() : JsonSerializer.SerializeToElement(new { }));
            else
                command.Completion.TrySetException(new OfficeWorkerException("jsa_command_failed", body.TryGetProperty("error", out var error) ? error.GetString() ?? "WPS JSA 操作失败" : "WPS JSA 操作失败"));
            await WriteJsonAsync(response, new { ok = true }, cancellationToken);
            return;
        }
        response.StatusCode = 404;
        response.Close();
    }

    private static async Task WriteJsonAsync(HttpListenerResponse response, object value, CancellationToken cancellationToken)
    {
        var bytes = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(value, new JsonSerializerOptions(JsonSerializerDefaults.Web)));
        response.StatusCode = 200;
        response.ContentType = "application/json; charset=utf-8";
        response.ContentLength64 = bytes.Length;
        await response.OutputStream.WriteAsync(bytes, cancellationToken);
        response.Close();
    }

    private sealed class PendingCommand(string id, string type, object? data)
    {
        public string Id { get; } = id;
        public string Type { get; } = type;
        public object? Data { get; } = data;
        public bool Delivered { get; set; }
        public TaskCompletionSource<JsonElement> Completion { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
    }
}
