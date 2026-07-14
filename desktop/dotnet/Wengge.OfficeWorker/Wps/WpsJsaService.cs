using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Wengge.OfficeWorker.Com;
using Wengge.OfficeWorker.Protocol;
using Wengge.OfficeWorker.Runtime;

namespace Wengge.OfficeWorker.Wps;

internal sealed class WpsJsaService : IDisposable
{
    private const string AddonName = "WenggeJsaBridge";
    private const string AddonDirectoryName = "WenggeJsaBridge_";
    private const string AddonVersion = "1";
    private readonly OfficeApplicationProvider applications;
    private readonly JsaHttpTransport transport = new();

    public WpsJsaService(OfficeApplicationProvider applications) => this.applications = applications;

    public async Task<object?> DetectAsync(CancellationToken cancellationToken)
    {
        using var active = applications.TryGetActive(["Ket.Application"]);
        if (active is null) return Unsupported("WPS JSA 仅在已运行的 WPS 表格中可用");
        var installation = ReadInstallation();
        if (!installation.Installed || string.IsNullOrWhiteSpace(installation.Token))
            return NotReady("WPS JSA 内部桥接尚未安装；首次写入时会自动安装");
        await transport.StartAsync(installation.Token, cancellationToken);
        if (!await transport.WaitForClientAsync(1_200, cancellationToken))
            return NotReady("WPS JSA 加载项尚未连接，请完全退出并重新打开 WPS 表格");
        try
        {
            await transport.SendAsync("detect", null, cancellationToken);
            return new { language = "javascript", supported = true, ready = true, @internal = true, engine = "WPS JSA" };
        }
        catch (Exception exception)
        {
            return NotReady(exception.Message);
        }
    }

    public async Task<object?> WriteAsync(System.Text.Json.JsonElement parameters, CancellationToken cancellationToken)
    {
        using var active = applications.TryGetActive(["Ket.Application"]);
        if (active is null) throw new OfficeWorkerException("wps_not_connected", "JavaScript 内部宏仅支持已运行的 WPS 表格");
        var code = NormalizeSource(parameters.RequiredString("code"));
        var entryPoint = parameters.OptionalString("entryPoint")?.Trim();
        if (code.Length == 0) throw new OfficeWorkerException("invalid_params", "WPS JSA 代码不能为空");
        if (!string.IsNullOrWhiteSpace(entryPoint) && !HasEntryPoint(code, entryPoint))
            throw new OfficeWorkerException("invalid_params", $"WPS JSA 代码中找不到入口函数: {entryPoint}");
        var sourceDir = parameters.RequiredString("sourceDir");
        var installation = EnsureInstalled(sourceDir);
        if (installation.Changed)
            throw new OfficeWorkerException("jsa_restart_required", "WPS JSA 内部桥接已安装，请完全退出并重新打开 WPS 表格后再次执行");
        await transport.StartAsync(installation.Token, cancellationToken);
        var result = await transport.SendAsync("write", new
        {
            code,
            entryPoint,
            save = parameters.OptionalBoolean("save"),
        }, cancellationToken);
        var source = result.TryGetProperty("source", out var sourceValue) ? NormalizeSource(sourceValue.GetString() ?? string.Empty) : string.Empty;
        if (!string.Equals(source, code, StringComparison.Ordinal))
            throw new OfficeWorkerException("jsa_verify_failed", "WPS JSA 源码回读不一致");
        return new
        {
            language = "javascript",
            componentName = StringProperty(result, "componentName"),
            lineCount = IntProperty(result, "lineCount"),
            sourceVerified = true,
            entryPoint,
            entryPointVerified = BoolProperty(result, "entryPointVerified", string.IsNullOrWhiteSpace(entryPoint)),
            saved = BoolProperty(result, "saved"),
            workbookName = StringProperty(result, "workbookName"),
            host = "wps",
        };
    }

    public void Dispose() => transport.Dispose();

    private static Installation ReadInstallation()
    {
        var directory = AddonDirectory();
        var tokenPath = Path.Combine(directory, "bridge-token.txt");
        var installed = File.Exists(Path.Combine(directory, "index.html")) && File.Exists(tokenPath);
        return new Installation(
            installed,
            installed && File.Exists(Path.Combine(directory, "bridge-version.txt")) && File.ReadAllText(Path.Combine(directory, "bridge-version.txt")).Trim() == AddonVersion,
            installed ? File.ReadAllText(tokenPath).Trim() : string.Empty,
            false);
    }

    private static Installation EnsureInstalled(string sourceDir)
    {
        var existing = ReadInstallation();
        if (existing.Current)
        {
            UpsertPublishManifest();
            return existing;
        }
        var fullSource = Path.GetFullPath(sourceDir);
        if (!File.Exists(Path.Combine(fullSource, "index.html")))
            throw new OfficeWorkerException("jsa_assets_missing", $"安装包缺少 WPS JSA 桥接资源: {fullSource}");
        var directory = AddonDirectory();
        Directory.CreateDirectory(directory);
        CopyDirectory(fullSource, directory);
        var token = string.IsNullOrWhiteSpace(existing.Token) ? Guid.NewGuid().ToString("N") : existing.Token;
        File.WriteAllText(Path.Combine(directory, "bridge-token.txt"), token, new UTF8Encoding(false));
        File.WriteAllText(Path.Combine(directory, "bridge-version.txt"), AddonVersion, new UTF8Encoding(false));
        File.WriteAllText(Path.Combine(directory, "bridge-config.js"), $"window.WENGGE_JSA_BRIDGE={{port:45221,token:{JsonSerializer.Serialize(token)}}};\n", new UTF8Encoding(false));
        UpsertPublishManifest();
        return new Installation(true, true, token, true);
    }

    private static void UpsertPublishManifest()
    {
        var root = Path.GetDirectoryName(AddonDirectory())!;
        Directory.CreateDirectory(root);
        var path = Path.Combine(root, "publish.xml");
        var document = File.Exists(path) ? XDocument.Load(path, LoadOptions.PreserveWhitespace) : new XDocument(new XElement("jsplugins"));
        var rootElement = document.Root;
        if (rootElement?.Name.LocalName != "jsplugins") throw new OfficeWorkerException("jsa_manifest_invalid", "WPS publish.xml 根节点不是 jsplugins");
        var plugin = rootElement.Elements("jsplugin").FirstOrDefault(node => (string?)node.Attribute("name") == AddonName);
        if (plugin is null)
        {
            plugin = new XElement("jsplugin");
            rootElement.Add(plugin);
        }
        plugin.SetAttributeValue("name", AddonName);
        plugin.SetAttributeValue("type", "et");
        plugin.SetAttributeValue("url", $"file://%AppData%/kingsoft/wps/jsaddons/{AddonDirectoryName}/index.html");
        plugin.SetAttributeValue("debug", string.Empty);
        plugin.SetAttributeValue("enable", "enable_dev");
        using var writer = System.Xml.XmlWriter.Create(path, new System.Xml.XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = true });
        document.Save(writer);
    }

    private static void CopyDirectory(string source, string destination)
    {
        foreach (var directory in Directory.EnumerateDirectories(source, "*", SearchOption.AllDirectories))
            Directory.CreateDirectory(Path.Combine(destination, Path.GetRelativePath(source, directory)));
        foreach (var file in Directory.EnumerateFiles(source, "*", SearchOption.AllDirectories))
        {
            var target = Path.Combine(destination, Path.GetRelativePath(source, file));
            Directory.CreateDirectory(Path.GetDirectoryName(target)!);
            File.Copy(file, target, overwrite: true);
        }
    }

    private static string AddonDirectory()
    {
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        if (string.IsNullOrWhiteSpace(appData)) throw new OfficeWorkerException("appdata_missing", "无法确定 Windows AppData 目录");
        return Path.Combine(appData, "kingsoft", "wps", "jsaddons", AddonDirectoryName);
    }

    private static object Unsupported(string reason) => new { language = "javascript", supported = false, ready = false, @internal = true, engine = "WPS JSA", reason };
    private static object NotReady(string reason) => new { language = "javascript", supported = true, ready = false, @internal = true, engine = "WPS JSA", reason };
    private static string NormalizeSource(string source) => source.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n').TrimStart().TrimEnd('\n');
    private static bool HasEntryPoint(string code, string entryPoint) => JavaScriptEntryRegex(entryPoint.Split('.').Last()).IsMatch(code);
    private static Regex JavaScriptEntryRegex(string name) => new($@"(?:^|\n)\s*(?:export\s+)?(?:(?:async\s+)?function\s+{Regex.Escape(name)}\s*\(|(?:const|let|var)\s+{Regex.Escape(name)}\s*=)", RegexOptions.Multiline);
    private static string? StringProperty(JsonElement value, string name) => value.TryGetProperty(name, out var property) && property.ValueKind == JsonValueKind.String ? property.GetString() : null;
    private static int IntProperty(JsonElement value, string name) => value.TryGetProperty(name, out var property) && property.TryGetInt32(out var result) ? result : 0;
    private static bool BoolProperty(JsonElement value, string name, bool fallback = false) => value.TryGetProperty(name, out var property) && property.ValueKind is JsonValueKind.True or JsonValueKind.False ? property.GetBoolean() : fallback;

    private sealed record Installation(bool Installed, bool Current, string Token, bool Changed);

}
