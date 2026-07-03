/**
 * PresentationComBridge — PowerPoint/WPS 演示 COM 桥接实现
 *
 * 与 Excel bridge 保持同一策略：主进程通过 PowerShell COM 自动化执行
 * 打开、创建、读取、编辑和保存操作。
 */

import type { PresentationBridge } from "../../contracts/office";
import { executePowerShell, psVar } from "../../../automation/powershell";
import { safeJsonParse } from "../../../automation/json";

const PRESENTATION_PROG_IDS = ["PowerPoint.Application", "Wpp.Application", "Kwpp.Application"];

export type PresentationHost = "powerpoint" | "wpp";

interface PresentationOpenResult {
  presentationName: string;
  createdApp: boolean;
  progId?: string;
  fullName?: string;
  version?: string;
}

function progIdsLiteral(): string {
  return "@(" + PRESENTATION_PROG_IDS.map((id) => `'${id}'`).join(", ") + ")";
}

function psStringLiteral(value?: string | null): string {
  if (!value) return "$null";
  return `'${value.replace(/'/g, "''")}'`;
}

function psNullableVar(name: string, value?: string | null): string {
  return value ? psVar(name, value) : `$${name} = $null`;
}

/** 获取 PPT COM 对象的 PowerShell 脚本（优先已运行的实例） */
function acquirePresentationAppScript(allowCreate = true, preferredProgId?: string | null): string {
  const createBlock = allowCreate ? `
if ($null -eq $app) {
  if ($preferredProgId) {
    try { $app = New-Object -ComObject $preferredProgId; $progId = $preferredProgId; $createdApp = $true } catch {}
  }
}
if ($null -eq $app) {
  foreach ($id in $progIds) {
    if ($id -eq $preferredProgId) { continue }
    try { $app = New-Object -ComObject $id; $progId = $id; $createdApp = $true; break } catch {}
  }
}
` : "";
  const missingMessage = allowCreate
    ? "未找到可用的 PowerPoint/WPS 演示 COM 应用"
    : "PowerPoint 或 WPS 演示未运行，请先打开文档";
  return `
$progIds = ${progIdsLiteral()}
$preferredProgId = ${psStringLiteral(preferredProgId)}
$app = $null; $progId = $null; $createdApp = $false
if ($preferredProgId) {
  try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject($preferredProgId); $progId = $preferredProgId } catch {}
}
if ($null -eq $app) {
  foreach ($id in $progIds) {
    if ($id -eq $preferredProgId) { continue }
    try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject($id); $progId = $id; break } catch {}
  }
}
${createBlock}
if ($null -eq $app) { throw '${missingMessage}' }
$app.Visible = $true
`;
}

function targetPresentationResolverScript(): string {
  return `
function Resolve-TargetPresentation($app, $targetPath) {
  if ($targetPath) {
    try {
      $normalizedTarget = [System.IO.Path]::GetFullPath([string]$targetPath)
      foreach ($candidate in $app.Presentations) {
        try {
          if ($candidate.FullName -and ([System.IO.Path]::GetFullPath([string]$candidate.FullName) -ieq $normalizedTarget)) {
            return $candidate
          }
        } catch {}
      }
    } catch {}
  }
  try { if ($null -ne $app.ActivePresentation) { return $app.ActivePresentation } } catch {}
  return $null
}`;
}

function slideTextShapesScript(): string {
  return `
function Get-ShapeTextInfo($shape) {
  $text = ''; $hasText = $false
  try { if ($shape.HasTextFrame -and $shape.TextFrame.HasText) { $text = [string]$shape.TextFrame.TextRange.Text; $hasText = $true } } catch {}
  [pscustomobject]@{ name = $shape.Name; index = $shape.Id; type = $shape.Type; hasText = $hasText; text = $text }
}`;
}

function slideLayoutResolverScript(): string {
  return `
function Normalize-LayoutName($value) {
  if ($null -eq $value) { return '' }
  return ([string]$value).ToLowerInvariant() -replace '[\\s_\\-]+', ''
}
function Get-LayoutAliases($layoutKey) {
  switch (Normalize-LayoutName $layoutKey) {
    'title' { return @('title','titleslide','onlytitle','标题','标题幻灯片') }
    'blank' { return @('blank','空白','空白幻灯片') }
    default { return @('titlebody','titleandcontent','titlecontent','titleandtext','标题和内容','标题与内容','标题和正文','标题与正文') }
  }
}
function Resolve-CustomSlideLayout($presentation, $layoutKey) {
  $aliases = Get-LayoutAliases $layoutKey | ForEach-Object { Normalize-LayoutName $_ }
  foreach ($customLayout in $presentation.SlideMaster.CustomLayouts) {
    if ($aliases -contains (Normalize-LayoutName $customLayout.Name)) { return $customLayout }
  }
  return $null
}`;
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

  async createPresentation(filePath: string): Promise<{ success: boolean; presentationName?: string; error?: string }> {
    try {
      const result = await executePowerShell(`
${psVar("_filePath", filePath)}
${acquirePresentationAppScript(true, this.cachedProgId)}
$pres = $null
try { $pres = $app.Presentations.Add(); $pres.SaveAs($_filePath); [pscustomobject]@{ presentationName = $pres.Name; fullName = $pres.FullName; createdApp = $createdApp; progId = $progId; version = $app.Version } | ConvertTo-Json -Compress }
catch { if ($createdApp -and $null -ne $app) { try { $app.Quit() } catch {} }; throw }`);
      const data = safeJsonParse<PresentationOpenResult>(result, "powershell", "创建 PowerPoint 演示文稿");
      this.ownsPresentationApp = this.ownsPresentationApp || data.createdApp;
      if (data.progId) this.cachedProgId = data.progId;
      this.activePresentationPath = data.fullName || filePath;
      this._connected = true;
      this._host = isWpsPresentationProgId(data.progId || "") ? "wpp" : "powerpoint";
      this._version = data.version || this._version;
      return { success: true, presentationName: data.presentationName };
    } catch (err: any) {
      return { success: false, error: `创建 PowerPoint 演示文稿失败: ${err.message}` };
    }
  }

  async inspectPresentation(): Promise<unknown> {
    return this.executePresOp("检查 PPT 演示文稿", (app, _progId, pres) => `
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
    return this.executePresOp("读取幻灯片", (_app, _progId, pres) => `
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
    return this.executePresOp("添加幻灯片", (_app, _progId, pres) => `
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
    return this.executePresOp("设置形状文本", (_app, _progId, pres) => `
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
    return this.executePresOp("替换 PPT 文本", (_app, _progId, pres) => `
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

    for (const progId of PRESENTATION_PROG_IDS) {
      try {
        const result = await executePowerShell(`
          try {
            $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
            $ver = $app.Version
            "OK|${progId}|$ver"
          } catch { "FAIL" }
        `);
        if (result.startsWith("OK|")) {
          const parts = result.split("|");
          this.cachedProgId = parts[1];
          this._connected = true;
          this._host = isWpsPresentationProgId(progId) ? "wpp" : "powerpoint";
          this._version = parts[2];
          return this.cachedProgId!;
        }
      } catch { /* 继续检测下一个 */ }
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
    try {
      const result = await executePowerShell(`
        $p = Get-Process -Name "POWERPNT" -ErrorAction SilentlyContinue
        $w = Get-Process -Name "wpp" -ErrorAction SilentlyContinue
        if (-not $w) { $w = Get-Process -Name "wps" -ErrorAction SilentlyContinue }
        if ($p) { "PPT" } elseif ($w) { "WPP" } else { "NONE" }
      `);
      const trimmed = result.trim();
      if (trimmed === "NONE" || !trimmed) return { running: false, availableHosts: [] };
      const host: PresentationHost = trimmed === "WPP" ? "wpp" : "powerpoint";
      return { running: true, availableHosts: [host] };
    } catch {
      return { running: false, availableHosts: [] };
    }
  }

  private async verifyComAvailable(hosts: PresentationHost[]): Promise<{
    available: boolean; host: PresentationHost; version?: string; presentationName?: string; progId?: string;
  }> {
    for (const host of hosts) {
      for (const progId of progIdsForHost(host)) {
        try {
          const result = await executePowerShell(`
          try {
            $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
            "OK|$($app.Version)|$(if ($app.ActivePresentation) { $app.ActivePresentation.Name } else { '' })"
          } catch { "FAIL" }
        `);
          if (result.startsWith("OK|")) {
            const p = result.split("|");
            return { available: true, host, version: p[1] || undefined, presentationName: p[2] || undefined, progId };
          }
        } catch { /* next */ }
      }
    }
    return { available: false, host: hosts[0] || "powerpoint" };
  }

  private async verifyDirectCom(): Promise<{
    available: boolean; host: PresentationHost; version?: string; presentationName?: string; progId?: string;
  }> {
    for (const progId of PRESENTATION_PROG_IDS) {
      try {
        const result = await executePowerShell(`
          try {
            $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
            "OK|$($app.Version)|$(if ($app.ActivePresentation) { $app.ActivePresentation.Name } else { '' })"
          } catch { "FAIL" }
        `);
        if (result.startsWith("OK|")) {
          const p = result.split("|");
          const host: PresentationHost = isWpsPresentationProgId(progId) ? "wpp" : "powerpoint";
          return { available: true, host, version: p[1] || undefined, presentationName: p[2] || undefined, progId };
        }
      } catch { /* next */ }
    }
    return { available: false, host: "powerpoint" };
  }
}

function isWpsPresentationProgId(progId: string): boolean {
  return progId.toLowerCase().includes("wpp");
}

function progIdsForHost(host: PresentationHost): string[] {
  return host === "wpp"
    ? ["Wpp.Application", "Kwpp.Application"]
    : ["PowerPoint.Application"];
}
