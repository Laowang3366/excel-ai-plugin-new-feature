import { randomUUID } from "crypto";
import { cpSync, existsSync, mkdirSync, readFileSync, writeFileSync } from "fs";
import { createServer, type IncomingMessage, type Server, type ServerResponse } from "http";
import path from "path";

import type {
  JsaWriteOptions,
  JsaWriteResult,
  MacroLanguageCapability,
  WpsJsaBridge as WpsJsaBridgeContract,
} from "../../contracts/excel";
import { executePowerShell, psVar } from "../../../automation/powershell";
import type { ExcelComBridge } from "./excelComBridge";

const BRIDGE_PORT = 45221;
const ADDON_NAME = "WenggeJsaBridge";
const ADDON_DIR_NAME = `${ADDON_NAME}_`;
const COMMAND_TIMEOUT_MS = 15_000;
const ADDON_VERSION = "1";

interface BridgeCommand {
  id: string;
  type: "detect" | "write";
  data?: unknown;
}

interface BridgeResponse {
  id: string;
  ok: boolean;
  result?: unknown;
  error?: string;
}

interface PendingCommand {
  command: BridgeCommand;
  delivered: boolean;
  resolve: (value: unknown) => void;
  reject: (reason: Error) => void;
  timer: ReturnType<typeof setTimeout>;
}

class LocalJsaTransport {
  private server: Server | null = null;
  private starting: Promise<void> | null = null;
  private pending: PendingCommand | null = null;
  private token = "";
  private lastClientPollAt = 0;

  async start(token: string): Promise<void> {
    this.token = token;
    if (this.server) return;
    if (this.starting) return this.starting;
    this.starting = new Promise<void>((resolve, reject) => {
      const server = createServer((request, response) => {
        void this.handleRequest(request, response);
      });
      server.once("error", reject);
      server.listen(BRIDGE_PORT, "127.0.0.1", () => {
        server.removeListener("error", reject);
        this.server = server;
        resolve();
      });
    }).finally(() => {
      this.starting = null;
    });
    return this.starting;
  }

  async waitForClient(timeoutMs = 1_200): Promise<boolean> {
    if (this.isClientConnected()) return true;
    const deadline = Date.now() + timeoutMs;
    while (Date.now() < deadline) {
      await new Promise((resolve) => setTimeout(resolve, 100));
      if (this.isClientConnected()) return true;
    }
    return false;
  }

  isClientConnected(): boolean {
    return Date.now() - this.lastClientPollAt < 2_000;
  }

  async send(type: BridgeCommand["type"], data?: unknown): Promise<unknown> {
    if (!this.server) throw new Error("WPS JSA 本地桥接服务尚未启动");
    if (this.pending) throw new Error("已有 WPS JSA 写入任务正在执行");
    if (!await this.waitForClient(2_000)) {
      throw new Error("WPS JSA 加载项未连接。若刚安装加载项，请完全退出并重新打开 WPS 表格");
    }

    return new Promise((resolve, reject) => {
      const command: BridgeCommand = { id: randomUUID(), type, data };
      const timer = setTimeout(() => {
        if (this.pending?.command.id === command.id) this.pending = null;
        reject(new Error("WPS JSA 内部宏操作超时"));
      }, COMMAND_TIMEOUT_MS);
      this.pending = { command, delivered: false, resolve, reject, timer };
    });
  }

  private async handleRequest(request: IncomingMessage, response: ServerResponse): Promise<void> {
    response.setHeader("Access-Control-Allow-Origin", "*");
    response.setHeader("Access-Control-Allow-Headers", "Content-Type, X-Wengge-Token");
    if (request.method === "OPTIONS") {
      response.writeHead(204).end();
      return;
    }
    if (request.headers["x-wengge-token"] !== this.token) {
      response.writeHead(403).end();
      return;
    }

    if (request.method === "GET" && request.url === "/command") {
      this.lastClientPollAt = Date.now();
      if (!this.pending || this.pending.delivered) {
        response.writeHead(204).end();
        return;
      }
      this.pending.delivered = true;
      this.writeJson(response, 200, this.pending.command);
      return;
    }

    if (request.method === "POST" && request.url === "/response") {
      try {
        const body = await readJsonBody(request) as BridgeResponse;
        const pending = this.pending;
        if (!pending || body.id !== pending.command.id) {
          this.writeJson(response, 409, { error: "响应与当前命令不匹配" });
          return;
        }
        clearTimeout(pending.timer);
        this.pending = null;
        if (body.ok) pending.resolve(body.result);
        else pending.reject(new Error(body.error || "WPS JSA 内部宏操作失败"));
        this.writeJson(response, 200, { ok: true });
      } catch (error) {
        this.writeJson(response, 400, { error: error instanceof Error ? error.message : String(error) });
      }
      return;
    }

    response.writeHead(404).end();
  }

  private writeJson(response: ServerResponse, status: number, data: unknown): void {
    response.writeHead(status, { "Content-Type": "application/json; charset=utf-8" });
    response.end(JSON.stringify(data));
  }
}

const sharedTransport = new LocalJsaTransport();

export class WpsJsaBridge implements WpsJsaBridgeContract {
  constructor(private readonly comBridge: ExcelComBridge) {}

  async detectCapabilities(): Promise<MacroLanguageCapability> {
    if (this.comBridge.host !== "wps") {
      return unsupported("WPS JSA 仅在 WPS 表格中可用");
    }
    const install = getAddonInstallation();
    if (!install.installed || !install.token) {
      return supportedButNotReady("WPS JSA 内部桥接尚未安装；首次调用 macro.write 时会自动安装");
    }
    try {
      await sharedTransport.start(install.token);
      if (!await sharedTransport.waitForClient()) {
        return supportedButNotReady("WPS JSA 加载项尚未连接，请重新打开 WPS 表格");
      }
      await sharedTransport.send("detect");
      return {
        language: "javascript",
        supported: true,
        ready: true,
        internal: true,
        engine: "WPS JSA",
      };
    } catch (error) {
      return supportedButNotReady(error instanceof Error ? error.message : String(error));
    }
  }

  async writeCode(code: string, options: JsaWriteOptions = {}): Promise<JsaWriteResult> {
    if (this.comBridge.host !== "wps") {
      throw new Error("JavaScript 内部宏仅支持 WPS JSA；Microsoft Excel 请使用 VBA");
    }
    const normalizedCode = normalizeSource(code);
    const entryPoint = options.entryPoint?.trim();
    if (!normalizedCode) throw new Error("WPS JSA 代码不能为空");
    if (entryPoint && !hasJavaScriptEntryPoint(normalizedCode, entryPoint)) {
      throw new Error(`WPS JSA 代码中找不到入口函数: ${entryPoint}`);
    }

    const installation = await ensureAddonInstalled();
    if (installation.changed) {
      throw new Error("WPS JSA 内部桥接已安装。请完全退出并重新打开 WPS 表格后再次执行 macro.write");
    }
    await sharedTransport.start(installation.token);
    const result = await sharedTransport.send("write", {
      code: normalizedCode,
      entryPoint,
      save: options.save === true,
    }) as Omit<JsaWriteResult, "language" | "host" | "sourceVerified"> & { source: string };
    if (normalizeSource(result.source) !== normalizedCode) {
      throw new Error("WPS JSA 源码回读不一致");
    }
    return {
      language: "javascript",
      componentName: result.componentName,
      lineCount: result.lineCount,
      sourceVerified: true,
      entryPoint,
      entryPointVerified: result.entryPointVerified,
      saved: result.saved,
      workbookName: result.workbookName,
      host: "wps",
    };
  }

}

function unsupported(reason: string): MacroLanguageCapability {
  return {
    language: "javascript",
    supported: false,
    ready: false,
    internal: true,
    engine: "WPS JSA",
    reason,
  };
}

function supportedButNotReady(reason: string): MacroLanguageCapability {
  return {
    language: "javascript",
    supported: true,
    ready: false,
    internal: true,
    engine: "WPS JSA",
    reason,
  };
}

function normalizeSource(code: string): string {
  return code.replace(/\r\n?/g, "\n").replace(/\n+$/g, "").trimStart();
}

function hasJavaScriptEntryPoint(code: string, entryPoint: string): boolean {
  const name = entryPoint.split(".").pop()?.trim();
  if (!name) return false;
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return new RegExp(`(?:^|\\n)\\s*(?:export\\s+)?(?:async\\s+)?function\\s+${escaped}\\s*\\(|(?:^|\\n)\\s*(?:export\\s+)?(?:const|let|var)\\s+${escaped}\\s*=`, "m").test(code);
}

function getAddonInstallation(): { installed: boolean; current: boolean; token?: string } {
  const addonDir = getAddonDir();
  const tokenPath = path.join(addonDir, "bridge-token.txt");
  const versionPath = path.join(addonDir, "bridge-version.txt");
  const installed = existsSync(path.join(addonDir, "index.html")) && existsSync(tokenPath);
  return {
    installed,
    current: installed && existsSync(versionPath) && readFileSync(versionPath, "utf8").trim() === ADDON_VERSION,
    token: installed ? readFileSync(tokenPath, "utf8").trim() : undefined,
  };
}

async function ensureAddonInstalled(): Promise<{ token: string; changed: boolean }> {
  const addonDir = getAddonDir();
  const existing = getAddonInstallation();
  const token = existing.token || randomUUID();
  const sourceDir = resolveAddonSourceDir();
  mkdirSync(addonDir, { recursive: true });
  cpSync(sourceDir, addonDir, { recursive: true, force: true });
  writeFileSync(path.join(addonDir, "bridge-token.txt"), token, "utf8");
  writeFileSync(path.join(addonDir, "bridge-version.txt"), ADDON_VERSION, "utf8");
  writeFileSync(
    path.join(addonDir, "bridge-config.js"),
    `window.WENGGE_JSA_BRIDGE={port:${BRIDGE_PORT},token:${JSON.stringify(token)}};\n`,
    "utf8",
  );
  const manifestChanged = await upsertPublishManifest();
  return { token, changed: !existing.current || manifestChanged };
}

function resolveAddonSourceDir(): string {
  const candidates = [
    path.join(process.cwd(), "public", "wps-jsa-bridge"),
    process.resourcesPath ? path.join(process.resourcesPath, "public", "wps-jsa-bridge") : "",
  ].filter(Boolean);
  const source = candidates.find((candidate) => existsSync(path.join(candidate, "index.html")));
  if (!source) throw new Error("安装包缺少 WPS JSA 内部桥接资源");
  return source;
}

function getAddonDir(): string {
  const appData = process.env.APPDATA;
  if (!appData) throw new Error("无法确定 Windows AppData 目录");
  return path.join(appData, "kingsoft", "wps", "jsaddons", ADDON_DIR_NAME);
}

async function upsertPublishManifest(): Promise<boolean> {
  const appData = process.env.APPDATA;
  if (!appData) throw new Error("无法确定 Windows AppData 目录");
  const root = path.join(appData, "kingsoft", "wps", "jsaddons");
  const publishPath = path.join(root, "publish.xml");
  const addonUrl = `file://%AppData%/kingsoft/wps/jsaddons/${ADDON_DIR_NAME}/index.html`;
  const result = await executePowerShell(`
    ${psVar("_root", root)}
    ${psVar("_publishPath", publishPath)}
    ${psVar("_addonName", ADDON_NAME)}
    ${psVar("_addonUrl", addonUrl)}
    [System.IO.Directory]::CreateDirectory($_root) | Out-Null
    $xml = New-Object System.Xml.XmlDocument
    if ([System.IO.File]::Exists($_publishPath)) {
      $xml.Load($_publishPath)
      if ($xml.DocumentElement.Name -ne "jsplugins") { throw "WPS publish.xml 根节点不是 jsplugins" }
    } else {
      $declaration = $xml.CreateXmlDeclaration("1.0", "UTF-8", $null)
      $xml.AppendChild($declaration) | Out-Null
      $xml.AppendChild($xml.CreateElement("jsplugins")) | Out-Null
    }
    $changed = $false
    $plugin = $null
    foreach ($node in @($xml.DocumentElement.SelectNodes("jsplugin"))) {
      if ($node.GetAttribute("name") -eq $_addonName) { $plugin = $node; break }
    }
    if ($null -eq $plugin) {
      $plugin = $xml.CreateElement("jsplugin")
      $xml.DocumentElement.AppendChild($plugin) | Out-Null
      $changed = $true
    }
    if ($plugin.GetAttribute("type") -ne "et" -or
        $plugin.GetAttribute("url") -ne $_addonUrl -or
        $plugin.GetAttribute("enable") -ne "enable_dev") { $changed = $true }
    $plugin.SetAttribute("name", $_addonName)
    $plugin.SetAttribute("type", "et")
    $plugin.SetAttribute("url", $_addonUrl)
    $plugin.SetAttribute("debug", "")
    $plugin.SetAttribute("enable", "enable_dev")
    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
    $settings.Indent = $true
    $writer = [System.Xml.XmlWriter]::Create($_publishPath, $settings)
    try { $xml.Save($writer) } finally { $writer.Dispose() }
    [PSCustomObject]@{ changed = $changed } | ConvertTo-Json -Compress
  `);
  const parsed = JSON.parse(result) as { changed?: boolean };
  return parsed.changed === true;
}

function readJsonBody(request: IncomingMessage): Promise<unknown> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    let size = 0;
    request.on("data", (chunk: Buffer) => {
      size += chunk.length;
      if (size > 2 * 1024 * 1024) {
        reject(new Error("WPS JSA 响应过大"));
        request.destroy();
        return;
      }
      chunks.push(chunk);
    });
    request.on("end", () => {
      try {
        resolve(JSON.parse(Buffer.concat(chunks).toString("utf8")));
      } catch {
        reject(new Error("WPS JSA 响应不是有效 JSON"));
      }
    });
    request.on("error", reject);
  });
}
