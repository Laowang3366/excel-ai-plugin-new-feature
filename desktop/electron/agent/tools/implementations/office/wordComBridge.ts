/**
 * WordComBridge — Word/WPS 文字 COM 桥接实现
 *
 * 与 Excel bridge 保持同一策略：主进程通过 PowerShell COM 自动化执行
 * 打开、创建、读取、编辑和保存操作。
 */

import type { WordDocumentBridge } from "../../contracts/office";
import { executePowerShell, psVar } from "../../../automation/powershell";
import { safeJsonParse } from "../../../automation/json";
import {
  buildAcquireOfficeAppScript,
  buildTargetOfficeFileResolverScript,
  detectOfficeProcess,
  findActiveOfficeComProgId,
  psNullableVar,
  verifyDirectOfficeCom,
  verifyOfficeComAvailable,
} from "./officeComPowerShell";

const WORD_PROG_IDS = ["Word.Application", "Kwps.Application", "Wps.Application"];

export type WordHost = "word" | "wps";

interface WordDocumentOpenResult {
  documentName: string;
  createdApp: boolean;
  progId?: string;
  fullName?: string;
  version?: string;
}

/** 获取 Word COM 对象的 PowerShell 脚本（优先已运行的实例） */
function acquireWordAppScript(allowCreate = true, preferredProgId?: string | null): string {
  const missingMessage = allowCreate
    ? "未找到可用的 Word/WPS 文字 COM 应用"
    : "Word 或 WPS 文字未运行，请先打开文档";
  return buildAcquireOfficeAppScript({
    progIds: WORD_PROG_IDS,
    allowCreate,
    preferredProgId,
    missingMessage,
  });
}

function targetDocumentResolverScript(): string {
  return buildTargetOfficeFileResolverScript({
    functionName: "Resolve-TargetWordDocument",
    collectionProperty: "Documents",
    activeProperty: "ActiveDocument",
  });
}

export class WordComBridge implements WordDocumentBridge {
  private ownsWordApp = false;
  private cachedProgId: string | null = null;
  private activeDocumentPath: string | null = null;
  private _connected = false;
  private _host: WordHost | "unknown" = "unknown";
  private _version?: string;

  /** 同步返回当前连接状态 */
  isConnected(): boolean {
    return this._connected;
  }

  /** 同步返回当前宿主类型 */
  getHost(): string {
    return this._host;
  }

  async detectStatus(): Promise<{
    connected: boolean; host: string; version?: string; documentName?: string;
  }> {
    try {
      const proc = await this.detectWordProcess();
      if (proc.running) {
        const comResult = await this.verifyComAvailable(proc.availableHosts);
        if (comResult.available) {
          this._connected = true; this._host = comResult.host; this._version = comResult.version;
          if (comResult.progId) this.cachedProgId = comResult.progId;
          return { connected: true, host: comResult.host, version: comResult.version, documentName: comResult.documentName };
        }
        this._connected = false;
        return { connected: false, host: proc.availableHosts[0] || "unknown" };
      }
      const directResult = await this.verifyDirectCom();
      if (directResult.available) {
        this._connected = true; this._host = directResult.host; this._version = directResult.version;
        if (directResult.progId) this.cachedProgId = directResult.progId;
        return { connected: true, host: directResult.host, version: directResult.version, documentName: directResult.documentName };
      }
      this._connected = false; this._host = "unknown";
      return { connected: false, host: "unknown" };
    } catch {
      this._connected = false;
      return { connected: false, host: "unknown" };
    }
  }

  async openDocument(filePath: string): Promise<{ success: boolean; documentName?: string; error?: string }> {
    try {
      const result = await executePowerShell(`
${psVar("_filePath", filePath)}
${acquireWordAppScript(true, this.cachedProgId)}
$doc = $null
try {
  $doc = $app.Documents.Open($_filePath); $doc.Activate()
  [pscustomobject]@{ documentName = $doc.Name; fullName = $doc.FullName; createdApp = $createdApp; progId = $progId; version = $app.Version } | ConvertTo-Json -Compress
} catch {
  if ($createdApp -and $null -ne $app) { try { $app.Quit() } catch {} }; throw
}`);
      const data = safeJsonParse<WordDocumentOpenResult>(result, "powershell", "打开 Word 文档");
      this.ownsWordApp = this.ownsWordApp || data.createdApp;
      if (data.progId) this.cachedProgId = data.progId;
      this.activeDocumentPath = data.fullName || filePath;
      this._connected = true;
      this._host = isWpsWordProgId(data.progId || "") ? "wps" : "word";
      this._version = data.version || this._version;
      return { success: true, documentName: data.documentName };
    } catch (err: any) {
      return { success: false, error: `打开 Word 文档失败: ${err.message}` };
    }
  }

  async createDocument(filePath: string): Promise<{ success: boolean; documentName?: string; error?: string }> {
    try {
      const result = await executePowerShell(`
${psVar("_filePath", filePath)}
${acquireWordAppScript(true, this.cachedProgId)}
$doc = $null
try {
  $doc = $app.Documents.Add(); $doc.SaveAs2($_filePath); $doc.Activate()
  [pscustomobject]@{ documentName = $doc.Name; fullName = $doc.FullName; createdApp = $createdApp; progId = $progId; version = $app.Version } | ConvertTo-Json -Compress
} catch {
  if ($createdApp -and $null -ne $app) { try { $app.Quit() } catch {} }; throw
}`);
      const data = safeJsonParse<WordDocumentOpenResult>(result, "powershell", "创建 Word 文档");
      this.ownsWordApp = this.ownsWordApp || data.createdApp;
      if (data.progId) this.cachedProgId = data.progId;
      this.activeDocumentPath = data.fullName || filePath;
      this._connected = true;
      this._host = isWpsWordProgId(data.progId || "") ? "wps" : "word";
      this._version = data.version || this._version;
      return { success: true, documentName: data.documentName };
    } catch (err: any) {
      return { success: false, error: `创建 Word 文档失败: ${err.message}` };
    }
  }

  async inspectDocument(): Promise<unknown> {
    return this.executeDocOp("检查 Word 文档", (app, _progId, doc) => `
$paragraphPreview = @()
$maxPreview = [Math]::Min(8, $doc.Paragraphs.Count)
for ($i = 1; $i -le $maxPreview; $i++) {
  $text = [string]$doc.Paragraphs.Item($i).Range.Text; $paragraphPreview += $text.Trim()
}
[pscustomobject]@{
  app = $app.Name; progId = $progId; name = $doc.Name; path = $doc.FullName
  paragraphs = $doc.Paragraphs.Count; tables = $doc.Tables.Count; sections = $doc.Sections.Count
  words = $doc.Words.Count; characters = $doc.Characters.Count; preview = $paragraphPreview
} | ConvertTo-Json -Depth 5 -Compress
`);
  }

  async readText(maxChars = 12000): Promise<unknown> {
    return this.executeDocOp("读取 Word 文本", (_app, _progId, doc) => `
$maxChars = ${Math.max(1, Math.floor(maxChars))}
$text = [string]$doc.Content.Text; $charCount = $text.Length; $truncated = $false
if ($text.Length -gt $maxChars) { $text = $text.Substring(0, $maxChars); $truncated = $true }
[pscustomobject]@{ name = $doc.Name; path = $doc.FullName; text = $text; charCount = $charCount; truncated = $truncated } | ConvertTo-Json -Depth 4 -Compress
`);
  }

  async insertText(text: string, position = "end"): Promise<unknown> {
    return this.executeDocOp("插入 Word 文本", (_app, _progId, doc) => `
${psVar("_text", text)}
${psVar("_position", position)}
switch ($_position) {
  'start' { $r = $doc.Range(0, 0); $r.InsertBefore($_text) }
  'selection' { $app.Selection.TypeText($_text) }
  default { $end = [Math]::Max(0, $doc.Content.End - 1); $r = $doc.Range($end, $end); $r.InsertAfter($_text) }
}
[pscustomobject]@{ inserted = $true; position = $_position; characters = $_text.Length } | ConvertTo-Json -Compress
`);
  }

  async insertHeading(text: string, level = 1, position = "end"): Promise<unknown> {
    const headingLevel = String(Math.min(9, Math.max(1, Math.floor(level))));
    return this.executeDocOp("插入 Word 标题", (_app, _progId, doc) => `
${psVar("_text", text)}
${psVar("_position", position)}
${psVar("_headingLevel", headingLevel)}
$headingLevel = [int]$_headingLevel
$line = $_text + [Environment]::NewLine; $start = 0
switch ($_position) {
  'start' { $r = $doc.Range(0, 0); $r.InsertBefore($line); $start = 0 }
  'selection' { $sr = $app.Selection.Range; $start = $sr.Start; $app.Selection.TypeText($line) }
  default { $start = [Math]::Max(0, $doc.Content.End - 1); $r = $doc.Range($start, $start); $r.InsertAfter($line) }
}
$hr = $doc.Range($start, $start + $_text.Length)
# wdStyleHeading1=-2 → HeadingN = -1 - N
$hr.Style = $doc.Styles.Item(-1 - $headingLevel)
[pscustomobject]@{ inserted = $true; position = $_position; level = $headingLevel; characters = $_text.Length } | ConvertTo-Json -Compress
`);
  }

  async replaceText(findText: string, replaceText: string, matchCase = false): Promise<unknown> {
    return this.executeDocOp("替换 Word 文本", (_app, _progId, doc) => `
${psVar("_findText", findText)}
${psVar("_replaceText", replaceText)}
$matchCase = ${matchCase ? "$true" : "$false"}
$count = 0; $r = $doc.Content
while ($true) {
  $f = $r.Find; $f.ClearFormatting(); $f.Replacement.ClearFormatting()
  $f.Text = $_findText; $f.Forward = $true; $f.Wrap = 0; $f.MatchCase = $matchCase
  if (-not $f.Execute()) { break }
  $count += 1; $r.Text = $_replaceText; $r = $doc.Range($r.End, $doc.Content.End)
}
[pscustomobject]@{ replacements = $count } | ConvertTo-Json -Compress
`);
  }

  async saveDocument(saveAsPath?: string): Promise<{ success: boolean; error?: string }> {
    const shouldQuitApp = this.ownsWordApp;
    try {
      await executePowerShell(`
${saveAsPath ? psVar("_saveAsPath", saveAsPath) : "$_saveAsPath = $null"}
$shouldQuitApp = ${shouldQuitApp ? "$true" : "$false"}
$doc = $null; $app = $null
try {
  ${acquireWordAppScript(false, this.cachedProgId)}
  ${psNullableVar("_activeDocumentPath", this.activeDocumentPath)}
  ${targetDocumentResolverScript()}
  $doc = Resolve-TargetWordDocument $app $_activeDocumentPath
  if ($null -eq $doc) { throw '当前没有活动 Word 文档' }
  if ($_saveAsPath) { $doc.SaveAs2($_saveAsPath) } else { $doc.Save() }
} finally {
  if ($shouldQuitApp) {
    if ($null -ne $doc) { try { $doc.Close($false) } catch {} }
    if ($null -ne $app) { try { $app.Quit() } catch {} }
  }
  if ($null -ne $doc) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) } catch {} }
  if ($null -ne $app) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch {} }
  [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
}`);
      if (saveAsPath) this.activeDocumentPath = saveAsPath;
      return { success: true };
    } catch (err: any) {
      return { success: false, error: `保存 Word 文档失败: ${err.message}` };
    } finally {
      if (shouldQuitApp) this.ownsWordApp = false;
    }
  }

  // ============================================================
  // 私有辅助：统一文档操作执行器（参考 Excel 的 ensureConnected 模式）
  // ============================================================

  /**
   * 统一文档操作执行：先 ensureConnected（获取 COM 连接），再执行具体脚本。
   * 与 Excel bridge 的 getWorkbookOperationDeps/ensureConnected 模式一致。
   */
  private async executeDocOp<T>(
    opName: string,
    buildScript: (app: string, progId: string, doc: string) => string,
  ): Promise<T> {
    const progId = await this.ensureConnected();
    try {
      const result = await executePowerShell(`
${acquireWordAppScript(false, progId)}
${psNullableVar("_activeDocumentPath", this.activeDocumentPath)}
${targetDocumentResolverScript()}
$doc = Resolve-TargetWordDocument $app $_activeDocumentPath
if ($null -eq $doc) { throw '当前没有活动 Word 文档，请先打开或创建文档' }
${buildScript("$app", `'${progId}'`, "$doc")}
`);
      const data = safeJsonParse<Record<string, unknown>>(result, "powershell", opName);
      if (typeof data.progId === "string") this.cachedProgId = data.progId;
      return data as unknown as T;
    } catch (err: any) {
      throw new Error(`${opName}失败: ${err.message}`);
    }
  }

  /**
   * 确保 Word COM 连接可用（参考 Excel 的 detectAndConnect）。
   * 如果已有缓存的 ProgID 且连接有效则直接返回，否则重新检测。
   */
  private async ensureConnected(): Promise<string> {
    if (this.cachedProgId) {
      return this.cachedProgId;
    }
    if (this._connected && (this._host === "word" || this._host === "wps")) {
      this.cachedProgId = progIdsForHost(this._host)[0];
      return this.cachedProgId;
    }

    const activeCom = await findActiveOfficeComProgId({
      progIds: WORD_PROG_IDS,
      hostForProgId: hostForWordProgId,
    });
    if (activeCom) {
      this.cachedProgId = activeCom.progId;
      this._connected = true;
      this._host = activeCom.host;
      this._version = activeCom.version;
      return this.cachedProgId;
    }

    this._connected = false;
    throw new Error("Word 或 WPS 文字未运行，请先打开文档");
  }

  // ============================================================
  // 检测方法
  // ============================================================

  private async detectWordProcess(): Promise<{
    running: boolean; availableHosts: WordHost[];
  }> {
    return detectOfficeProcess<WordHost>({
      checks: [
        { token: "WORD", host: "word", processNames: ["WINWORD"] },
        { token: "WPS", host: "wps", processNames: ["wps"] },
      ],
    });
  }

  private async verifyComAvailable(hosts: WordHost[]): Promise<{
    available: boolean; host: WordHost; version?: string; documentName?: string; progId?: string;
  }> {
    const result = await verifyOfficeComAvailable({
      hosts,
      defaultHost: "word",
      progIdsForHost,
      activeObjectExpression: "$app.ActiveDocument",
    });
    return {
      available: result.available,
      host: result.host,
      version: result.version,
      documentName: result.activeName,
      progId: result.progId,
    };
  }

  private async verifyDirectCom(): Promise<{
    available: boolean; host: WordHost; version?: string; documentName?: string; progId?: string;
  }> {
    const result = await verifyDirectOfficeCom({
      progIds: WORD_PROG_IDS,
      defaultHost: "word",
      hostForProgId: hostForWordProgId,
      activeObjectExpression: "$app.ActiveDocument",
    });
    return {
      available: result.available,
      host: result.host,
      version: result.version,
      documentName: result.activeName,
      progId: result.progId,
    };
  }
}

function progIdsForHost(host: WordHost): string[] {
  return host === "wps" ? ["Kwps.Application", "Wps.Application"] : ["Word.Application"];
}

function isWpsWordProgId(progId: string): boolean {
  return progId.toLowerCase().includes("wps");
}

function hostForWordProgId(progId: string): WordHost {
  return isWpsWordProgId(progId) ? "wps" : "word";
}
