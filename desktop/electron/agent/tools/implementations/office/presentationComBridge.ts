/**
 * PresentationComBridge — PowerPoint/WPS 演示 COM 桥接实现
 *
 * 与 Excel bridge 保持同一策略：主进程通过 PowerShell COM 自动化执行
 * 打开、创建、读取、编辑和保存操作。
 */

import type { PresentationBridge } from "../../contracts/office";
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
import { slideLayoutResolverScript, slideTextShapesScript } from "./presentationComScripts";

const PRESENTATION_PROG_IDS = ["PowerPoint.Application", "Wpp.Application", "Kwpp.Application"];

export type PresentationHost = "powerpoint" | "wpp";

interface PresentationOpenResult {
  presentationName: string;
  createdApp: boolean;
  progId?: string;
  fullName?: string;
  version?: string;
}

/** 获取 PPT COM 对象的 PowerShell 脚本（优先已运行的实例） */
function acquirePresentationAppScript(allowCreate = true, preferredProgId?: string | null): string {
  const missingMessage = allowCreate
    ? "未找到可用的 PowerPoint/WPS 演示 COM 应用"
    : "PowerPoint 或 WPS 演示未运行，请先打开文档";
  return buildAcquireOfficeAppScript({
    progIds: PRESENTATION_PROG_IDS,
    allowCreate,
    preferredProgId,
    missingMessage,
  });
}

function targetPresentationResolverScript(): string {
  return buildTargetOfficeFileResolverScript({
    functionName: "Resolve-TargetPresentation",
    collectionProperty: "Presentations",
    activeProperty: "ActivePresentation",
  });
}

export class PresentationComBridge implements PresentationBridge {
  private ownsPresentationApp = false;
  private cachedProgId: string | null = null;
  private activePresentationPath: string | null = null;
  private _connected = false;
  private _host: PresentationHost | "unknown" = "unknown";
  private _version?: string;

  isConnected(): boolean { return this._connected; }
  getHost(): string { return this._host; }

  async detectStatus(): Promise<{
    connected: boolean; host: string; version?: string; presentationName?: string;
  }> {
    try {
      const proc = await this.detectPresentationProcess();
      if (proc.running) {
        const comResult = await this.verifyComAvailable(proc.availableHosts);
        if (comResult.available) {
          this._connected = true; this._host = comResult.host; this._version = comResult.version;
          if (comResult.progId) this.cachedProgId = comResult.progId;
          return { connected: true, host: comResult.host, version: comResult.version, presentationName: comResult.presentationName };
        }
        this._connected = false;
        return { connected: false, host: proc.availableHosts[0] || "unknown" };
      }
      const directResult = await this.verifyDirectCom();
      if (directResult.available) {
        this._connected = true; this._host = directResult.host; this._version = directResult.version;
        if (directResult.progId) this.cachedProgId = directResult.progId;
        return { connected: true, host: directResult.host, version: directResult.version, presentationName: directResult.presentationName };
      }
      this._connected = false; this._host = "unknown";
      return { connected: false, host: "unknown" };
    } catch {
      this._connected = false;
      return { connected: false, host: "unknown" };
    }
  }

  async openPresentation(filePath: string): Promise<{ success: boolean; presentationName?: string; error?: string }> {
    try {
      const result = await executePowerShell(`
${psVar("_filePath", filePath)}
${acquirePresentationAppScript(true, this.cachedProgId)}
$pres = $null
try { $pres = $app.Presentations.Open($_filePath); [pscustomobject]@{ presentationName = $pres.Name; fullName = $pres.FullName; createdApp = $createdApp; progId = $progId; version = $app.Version } | ConvertTo-Json -Compress }
catch { if ($createdApp -and $null -ne $app) { try { $app.Quit() } catch {} }; throw }`);
      const data = safeJsonParse<PresentationOpenResult>(result, "powershell", "打开 PowerPoint 演示文稿");
      this.ownsPresentationApp = this.ownsPresentationApp || data.createdApp;
      if (data.progId) this.cachedProgId = data.progId;
      this.activePresentationPath = data.fullName || filePath;
      this._connected = true;
      this._host = isWpsPresentationProgId(data.progId || "") ? "wpp" : "powerpoint";
      this._version = data.version || this._version;
      return { success: true, presentationName: data.presentationName };
    } catch (err: any) {
      return { success: false, error: `打开 PowerPoint 演示文稿失败: ${err.message}` };
    }
  }

  async inspectPresentation(): Promise<unknown> {
    return this.executePresOp("检查 PPT 演示文稿", (_app, _progId, _pres) => `
${slideTextShapesScript()}
$slides = @()
foreach ($slide in $pres.Slides) {
  $texts = @()
  foreach ($shape in $slide.Shapes) { $info = Get-ShapeTextInfo $shape; if ($info.hasText) { $texts += $info } }
  $slides += [pscustomobject]@{ index = $slide.SlideIndex; name = $slide.Name; shapeCount = $slide.Shapes.Count; textShapes = $texts }
}
[pscustomobject]@{ app = $app.Name; progId = $progId; name = $pres.Name; path = $pres.FullName; slideCount = $pres.Slides.Count; slides = $slides } | ConvertTo-Json -Depth 8 -Compress
`);
  }

  async readSlide(slideIndex: number): Promise<unknown> {
    const idx = Math.max(1, Math.floor(slideIndex));
    return this.executePresOp("读取幻灯片", (_app, _progId, _pres) => `
${slideTextShapesScript()}
$slideIndex = ${idx}
if ($slideIndex -gt $pres.Slides.Count) { throw "幻灯片序号超出范围: $slideIndex" }
$slide = $pres.Slides.Item($slideIndex)
$texts = @()
foreach ($shape in $slide.Shapes) { $texts += Get-ShapeTextInfo $shape }
[pscustomobject]@{ presentationName = $pres.Name; slideIndex = $slide.SlideIndex; name = $slide.Name; shapes = $texts } | ConvertTo-Json -Depth 6 -Compress
`);
  }

  async addSlide(title?: string, body?: string, layout = "title_body"): Promise<unknown> {
    return this.executePresOp("添加幻灯片", (_app, _progId, _pres) => `
${psVar("_title", title || "")}
${psVar("_body", body || "")}
${psVar("_layout", layout)}
${slideLayoutResolverScript()}
# PpSlideLayout: ppLayoutTitle=1, ppLayoutText=2, ppLayoutBlank=12
$layoutMap = @{ title = 1; title_body = 2; blank = 12 }
$layoutValue = if ($layoutMap.ContainsKey($_layout)) { $layoutMap[$_layout] } else { 2 }
$customLayout = Resolve-CustomSlideLayout $pres $_layout
if ($null -ne $customLayout) { $slide = $pres.Slides.AddSlide($pres.Slides.Count + 1, $customLayout) }
else { $slide = $pres.Slides.Add($pres.Slides.Count + 1, $layoutValue) }
if ($_title) {
  try { $slide.Shapes.Title.TextFrame.TextRange.Text = $_title } catch { $box = $slide.Shapes.AddTextbox(1, 40, 40, 640, 60); $box.TextFrame.TextRange.Text = $_title }
}
if ($_body) {
  $bodyShape = $null
  foreach ($shape in $slide.Shapes) {
    try { if ($shape.HasTextFrame -and $shape.TextFrame.HasText -and $shape.Name -ne $slide.Shapes.Title.Name) { $bodyShape = $shape; break } } catch {}
  }
  if ($null -eq $bodyShape) { $bodyShape = $slide.Shapes.AddTextbox(1, 60, 130, 620, 300) }
  $bodyShape.TextFrame.TextRange.Text = $_body
}
[pscustomobject]@{ slideIndex = $slide.SlideIndex; name = $slide.Name } | ConvertTo-Json -Compress
`);
  }

  async setShapeText(slideIndex: number, text: string, shapeName?: string, shapeIndex?: number): Promise<unknown> {
    const idx = Math.max(1, Math.floor(slideIndex));
    const sidx = shapeIndex ? Math.max(1, Math.floor(shapeIndex)) : 0;
    return this.executePresOp("设置形状文本", (_app, _progId, _pres) => `
${psVar("_text", text)}
${psVar("_shapeName", shapeName || "")}
$slideIndex = ${idx}; $shapeIndex = ${sidx}
if ($slideIndex -gt $pres.Slides.Count) { throw "幻灯片序号超出范围: $slideIndex" }
$slide = $pres.Slides.Item($slideIndex)
$target = $null
if ($_shapeName) { try { $target = $slide.Shapes.Item($_shapeName) } catch {} }
if ($null -eq $target -and $shapeIndex -gt 0) { try { $target = $slide.Shapes.Item($shapeIndex) } catch {} }
if ($null -eq $target) { foreach ($sh in $slide.Shapes) { try { if ($sh.HasTextFrame) { $target = $sh; break } } catch {} } }
if ($null -eq $target) { $target = $slide.Shapes.AddTextbox(1, 60, 120, 620, 300) }
$target.TextFrame.TextRange.Text = $_text
[pscustomobject]@{ slideIndex = $slide.SlideIndex; shapeName = $target.Name; textLength = $_text.Length } | ConvertTo-Json -Compress
`);
  }

  async replaceText(findText: string, replaceText: string, matchCase = false): Promise<unknown> {
    return this.executePresOp("替换 PPT 文本", (_app, _progId, _pres) => `
${psVar("_findText", findText)}
${psVar("_replaceText", replaceText)}
$matchCase = ${matchCase ? "$true" : "$false"}
if ([string]::IsNullOrEmpty($_findText)) { throw 'findText 不能为空' }
$comparison = if ($matchCase) { [StringComparison]::Ordinal } else { [StringComparison]::OrdinalIgnoreCase }
$replacements = 0; $changedShapes = @()
foreach ($slide in $pres.Slides) {
  foreach ($shape in $slide.Shapes) {
    try {
      if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
        $range = $shape.TextFrame.TextRange; $oldText = [string]$range.Text
        $sb = [System.Text.StringBuilder]::new(); $cursor = 0; $local = 0
        while ($cursor -lt $oldText.Length) {
          $foundAt = $oldText.IndexOf($_findText, $cursor, $comparison)
          if ($foundAt -lt 0) { break }
          [void]$sb.Append($oldText.Substring($cursor, $foundAt - $cursor))
          [void]$sb.Append($_replaceText)
          $cursor = $foundAt + $_findText.Length; $local += 1
        }
        if ($local -gt 0) {
          [void]$sb.Append($oldText.Substring($cursor)); $range.Text = $sb.ToString()
          $replacements += $local
          $changedShapes += [pscustomobject]@{ slideIndex = $slide.SlideIndex; shapeName = $shape.Name; replacements = $local }
        }
      }
    } catch {}
  }
}
[pscustomobject]@{ replacements = $replacements; changedShapes = $changedShapes } | ConvertTo-Json -Depth 5 -Compress
`);
  }

  async savePresentation(saveAsPath?: string): Promise<{ success: boolean; error?: string }> {
    const shouldQuitApp = this.ownsPresentationApp;
    try {
      await executePowerShell(`
${saveAsPath ? psVar("_saveAsPath", saveAsPath) : "$_saveAsPath = $null"}
$shouldQuitApp = ${shouldQuitApp ? "$true" : "$false"}
$pres = $null; $app = $null
try {
  ${acquirePresentationAppScript(false, this.cachedProgId)}
  ${psNullableVar("_activePresentationPath", this.activePresentationPath)}
  ${targetPresentationResolverScript()}
  $pres = Resolve-TargetPresentation $app $_activePresentationPath
  if ($null -eq $pres) { throw '当前没有活动 PowerPoint 演示文稿' }
  if ($_saveAsPath) { $pres.SaveAs($_saveAsPath) } else { $pres.Save() }
} finally {
  if ($shouldQuitApp) {
    if ($null -ne $pres) { try { $pres.Close() } catch {} }
    if ($null -ne $app) { try { $app.Quit() } catch {} }
  }
  if ($null -ne $pres) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) } catch {} }
  if ($null -ne $app) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch {} }
  [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
}`);
      if (saveAsPath) this.activePresentationPath = saveAsPath;
      return { success: true };
    } catch (err: any) {
      return { success: false, error: `保存 PowerPoint 演示文稿失败: ${err.message}` };
    } finally {
      if (shouldQuitApp) this.ownsPresentationApp = false;
    }
  }

  // ============================================================
  // 私有辅助：统一演示操作执行器
  // ============================================================

  private async executePresOp<T>(
    opName: string,
    buildScript: (app: string, progId: string, pres: string) => string,
  ): Promise<T> {
    const progId = await this.ensureConnected();
    try {
      const result = await executePowerShell(`
${acquirePresentationAppScript(false, progId)}
${psNullableVar("_activePresentationPath", this.activePresentationPath)}
${targetPresentationResolverScript()}
$pres = Resolve-TargetPresentation $app $_activePresentationPath
if ($null -eq $pres) { throw '当前没有活动 PowerPoint 演示文稿，请先打开或创建文档' }
${buildScript("$app", `'${progId}'`, "$pres")}
`);
      const data = safeJsonParse<Record<string, unknown>>(result, "powershell", opName);
      if (typeof data.progId === "string") this.cachedProgId = data.progId;
      return data as unknown as T;
    } catch (err: any) {
      throw new Error(`${opName}失败: ${err.message}`);
    }
  }

  private async ensureConnected(): Promise<string> {
    if (this.cachedProgId) {
      return this.cachedProgId;
    }
    if (this._connected && (this._host === "powerpoint" || this._host === "wpp")) {
      this.cachedProgId = progIdsForHost(this._host)[0];
      return this.cachedProgId;
    }

    const activeCom = await findActiveOfficeComProgId({
      progIds: PRESENTATION_PROG_IDS,
      hostForProgId: hostForPresentationProgId,
    });
    if (activeCom) {
      this.cachedProgId = activeCom.progId;
      this._connected = true;
      this._host = activeCom.host;
      this._version = activeCom.version;
      return this.cachedProgId;
    }

    this._connected = false;
    throw new Error("PowerPoint 或 WPS 演示未运行，请先打开文档");
  }

  // ============================================================
  // 检测方法
  // ============================================================

  private async detectPresentationProcess(): Promise<{
    running: boolean; availableHosts: PresentationHost[];
  }> {
    return detectOfficeProcess<PresentationHost>({
      checks: [
        { token: "PPT", host: "powerpoint", processNames: ["POWERPNT"] },
        { token: "WPP", host: "wpp", processNames: ["wpp", "wps"] },
      ],
    });
  }

  private async verifyComAvailable(hosts: PresentationHost[]): Promise<{
    available: boolean; host: PresentationHost; version?: string; presentationName?: string; progId?: string;
  }> {
    const result = await verifyOfficeComAvailable({
      hosts,
      defaultHost: "powerpoint",
      progIdsForHost,
      activeObjectExpression: "$app.ActivePresentation",
    });
    return {
      available: result.available,
      host: result.host,
      version: result.version,
      presentationName: result.activeName,
      progId: result.progId,
    };
  }

  private async verifyDirectCom(): Promise<{
    available: boolean; host: PresentationHost; version?: string; presentationName?: string; progId?: string;
  }> {
    const result = await verifyDirectOfficeCom({
      progIds: PRESENTATION_PROG_IDS,
      defaultHost: "powerpoint",
      hostForProgId: hostForPresentationProgId,
      activeObjectExpression: "$app.ActivePresentation",
    });
    return {
      available: result.available,
      host: result.host,
      version: result.version,
      presentationName: result.activeName,
      progId: result.progId,
    };
  }
}

function isWpsPresentationProgId(progId: string): boolean {
  return progId.toLowerCase().includes("wpp");
}

function hostForPresentationProgId(progId: string): PresentationHost {
  return isWpsPresentationProgId(progId) ? "wpp" : "powerpoint";
}

function progIdsForHost(host: PresentationHost): string[] {
  return host === "wpp"
    ? ["Wpp.Application", "Kwpp.Application"]
    : ["PowerPoint.Application"];
}
