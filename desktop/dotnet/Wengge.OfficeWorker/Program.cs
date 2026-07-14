using System.Text;
using Wengge.OfficeWorker.Protocol;
using Wengge.OfficeWorker.Runtime;

Console.InputEncoding = Encoding.UTF8;
Console.OutputEncoding = new UTF8Encoding(false);

using var worker = OfficeWorkerHost.Create();
var server = new JsonRpcServer(worker.DispatchAsync, Console.In, Console.Out, Console.Error);
await server.RunAsync(CancellationToken.None);
