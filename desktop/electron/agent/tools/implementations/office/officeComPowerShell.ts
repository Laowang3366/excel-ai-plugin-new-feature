import { executePowerShell, psVar } from "../../../automation/powershell";
import { officeInstanceDiscoveryScript } from "./officeDocumentComBridge";

export interface AcquireOfficeAppScriptOptions {
  progIds: readonly string[];
  allowCreate?: boolean;
  reuseAnyActive?: boolean;
  preferredProgId?: string | null;
  missingMessage: string;
  visible?: boolean | number;
  appKind?: "excel" | "word" | "presentation";
  targetPathExpression?: string;
  instanceIdExpression?: string;
}

export interface CreateOfficeAppScriptOptions {
  progIds: readonly string[];
  missingMessage: string;
  visible?: boolean | number;
}

export interface TargetOfficeFileResolverScriptOptions {
  functionName: string;
  collectionProperty: string;
  activeProperty: string;
}

export interface OfficeComVerificationResult<THost extends string> {
  available: boolean;
  host: THost;
  version?: string;
  activeName?: string;
  progId?: string;
}

export interface VerifyOfficeComAvailableOptions<THost extends string> {
  hosts: THost[];
  defaultHost: THost;
  progIdsForHost: (host: THost) => readonly string[];
  activeObjectExpression: string;
}

export interface VerifyDirectOfficeComOptions<THost extends string> {
  progIds: readonly string[];
  defaultHost: THost;
  hostForProgId: (progId: string) => THost;
  activeObjectExpression: string;
}

export interface FindActiveOfficeComProgIdOptions<THost extends string> {
  progIds: readonly string[];
  hostForProgId: (progId: string) => THost;
}

export interface ActiveOfficeComProgId<THost extends string> {
  progId: string;
  host: THost;
  version?: string;
}

export interface DetectOfficeProcessCheck<THost extends string> {
  token: string;
  host: THost;
  processNames: readonly string[];
}

export interface DetectOfficeProcessOptions<THost extends string> {
  checks: readonly DetectOfficeProcessCheck<THost>[];
}

export interface DetectOfficeProcessResult<THost extends string> {
  running: boolean;
  availableHosts: THost[];
}

export function progIdsLiteral(progIds: readonly string[]): string {
  return "@(" + progIds.map((id) => `'${id}'`).join(", ") + ")";
}

export function psStringLiteral(value?: string | null): string {
  if (!value) return "$null";
  return `'${value.replace(/'/g, "''")}'`;
}

export function psNullableVar(name: string, value?: string | null): string {
  return value ? psVar(name, value) : `$${name} = $null`;
}

export async function detectOfficeProcess<THost extends string>(
  options: DetectOfficeProcessOptions<THost>,
): Promise<DetectOfficeProcessResult<THost>> {
  try {
    const script = buildOfficeProcessDetectionScript(options.checks);
    const result = await executePowerShell(script);
    const token = result.trim();
    const match = options.checks.find((check) => check.token === token);
    return match
      ? { running: true, availableHosts: [match.host] }
      : { running: false, availableHosts: [] };
  } catch {
    return { running: false, availableHosts: [] };
  }
}

export function buildAcquireOfficeAppScript(options: AcquireOfficeAppScriptOptions): string {
  const {
    progIds,
    allowCreate = true,
    reuseAnyActive = true,
    preferredProgId,
    missingMessage,
    visible = true,
    appKind,
    targetPathExpression = "$_filePath",
    instanceIdExpression = "$actionParams.instanceId",
  } = options;
  const instanceBlock = appKind ? `
${officeInstanceDiscoveryScript()}
$actionTargetPath = [string](${targetPathExpression})
$actionInstanceId = [string](${instanceIdExpression})
if ($actionTargetPath -and $actionInstanceId) {
  $actionWantedPath = [IO.Path]::GetFullPath($actionTargetPath)
  $actionCandidates = @(Get-AllOfficeDocumentHandles '${appKind}' | Where-Object {
    $pathMatches = try { [IO.Path]::GetFullPath([string]$_.document.FullName) -ieq $actionWantedPath } catch { $false }
    $instanceMatches = -not $actionInstanceId -or [string]$_.instanceId -eq $actionInstanceId
    $pathMatches -and $instanceMatches
  })
  if ($actionCandidates.Count -gt 1) { throw '找到多个 Office 文档候选，请传 office.documents.list 返回的 instanceId' }
  if ($actionCandidates.Count -eq 1) { $app = $actionCandidates[0].application; $progId = [string]$actionCandidates[0].progId }
}
` : "";
  const createBlock = allowCreate ? `
if ($null -eq $app) {
  if ($preferredProgId) {
    foreach ($attempt in 1..3) {
      $candidate = $null
      try {
        $candidate = New-Object -ComObject $preferredProgId
        $null = [string]$candidate.Version
        $app = $candidate; $candidate = $null; $progId = $preferredProgId; $createdApp = $true; break
      } catch {
        $lastCreateError = $_
        if ($null -ne $candidate) { try { $candidate.Quit() } catch {}; try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($candidate) } catch {} }
      }
      Start-Sleep -Milliseconds (200 * $attempt)
    }
  }
}
if ($null -eq $app) {
  foreach ($id in $progIds) {
    if ($id -eq $preferredProgId) { continue }
    foreach ($attempt in 1..3) {
      $candidate = $null
      try {
        $candidate = New-Object -ComObject $id
        $null = [string]$candidate.Version
        $app = $candidate; $candidate = $null; $progId = $id; $createdApp = $true; break
      } catch {
        $lastCreateError = $_
        if ($null -ne $candidate) { try { $candidate.Quit() } catch {}; try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($candidate) } catch {} }
      }
      Start-Sleep -Milliseconds (200 * $attempt)
    }
    if ($null -ne $app) { break }
  }
}
` : "";
  const activeBlock = reuseAnyActive ? `
if ($null -eq $app -and $preferredProgId) {
  try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject($preferredProgId); $progId = $preferredProgId } catch {}
}
if ($null -eq $app) {
  foreach ($id in $progIds) {
    if ($id -eq $preferredProgId) { continue }
    try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject($id); $progId = $id; break } catch {}
  }
}
` : "";
  return `
$progIds = ${progIdsLiteral(progIds)}
$preferredProgId = ${psStringLiteral(preferredProgId)}
$app = $null; $progId = $null; $createdApp = $false; $lastCreateError = $null
${instanceBlock}
${activeBlock}
${createBlock}
if ($null -eq $app) {
  $detail = if ($null -ne $lastCreateError) { ': ' + [string]$lastCreateError.Exception.Message } else { '' }
  throw '${missingMessage}' + $detail
}
$app.Visible = ${powerShellScalar(visible)}
`;
}

export function buildCreateOfficeAppScript(options: CreateOfficeAppScriptOptions): string {
  const visible = powerShellScalar(options.visible ?? true);
  return `
$app = $null; $progId = $null; $createdApp = $false
foreach ($id in ${progIdsLiteral(options.progIds)}) {
  try { $app = New-Object -ComObject $id; $progId = $id; $createdApp = $true; break } catch {}
}
if ($null -eq $app) { throw '${options.missingMessage.replace(/'/g, "''")}' }
$app.Visible = ${visible}
try { $app.DisplayAlerts = $false } catch {}
`;
}

function powerShellScalar(value: boolean | number): string {
  return typeof value === "number" ? String(value) : value ? "$true" : "$false";
}

export function buildTargetOfficeFileResolverScript(options: TargetOfficeFileResolverScriptOptions): string {
  const { functionName, collectionProperty, activeProperty } = options;
  return `
function ${functionName}($app, $targetPath) {
  if ($targetPath) {
    try {
      $normalizedTarget = [System.IO.Path]::GetFullPath([string]$targetPath)
      foreach ($candidate in $app.${collectionProperty}) {
        try {
          if ($candidate.FullName -and ([System.IO.Path]::GetFullPath([string]$candidate.FullName) -ieq $normalizedTarget)) {
            return $candidate
          }
        } catch {}
      }
    } catch {}
    return $null
  }
  try { if ($null -ne $app.${activeProperty}) { return $app.${activeProperty} } } catch {}
  return $null
}`;
}

export async function verifyOfficeComAvailable<THost extends string>(
  options: VerifyOfficeComAvailableOptions<THost>,
): Promise<OfficeComVerificationResult<THost>> {
  const { hosts, defaultHost, progIdsForHost, activeObjectExpression } = options;
  for (const host of hosts) {
    for (const progId of progIdsForHost(host)) {
      const result = await verifySingleOfficeComProgId(progId, activeObjectExpression);
      if (result) {
        return {
          available: true,
          host,
          version: result.version,
          activeName: result.activeName,
          progId,
        };
      }
    }
  }
  return { available: false, host: hosts[0] || defaultHost };
}

export async function verifyDirectOfficeCom<THost extends string>(
  options: VerifyDirectOfficeComOptions<THost>,
): Promise<OfficeComVerificationResult<THost>> {
  const { progIds, defaultHost, hostForProgId, activeObjectExpression } = options;
  for (const progId of progIds) {
    const result = await verifySingleOfficeComProgId(progId, activeObjectExpression);
    if (result) {
      return {
        available: true,
        host: hostForProgId(progId),
        version: result.version,
        activeName: result.activeName,
        progId,
      };
    }
  }
  return { available: false, host: defaultHost };
}

export async function findActiveOfficeComProgId<THost extends string>(
  options: FindActiveOfficeComProgIdOptions<THost>,
): Promise<ActiveOfficeComProgId<THost> | null> {
  const { progIds, hostForProgId } = options;
  for (const progId of progIds) {
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
        return {
          progId: parts[1],
          host: hostForProgId(progId),
          version: parts[2] || undefined,
        };
      }
    } catch { /* next */ }
  }
  return null;
}

function buildOfficeProcessDetectionScript<THost extends string>(
  checks: readonly DetectOfficeProcessCheck<THost>[],
): string {
  const declarations = checks.map((check, index) => `
$match${index} = $false
foreach ($name in ${progIdsLiteral(check.processNames)}) {
  if (Get-Process -Name $name -ErrorAction SilentlyContinue) { $match${index} = $true; break }
}`).join("\n");
  const branches = checks
    .map((check, index) => `${index === 0 ? "if" : "elseif"} ($match${index}) { "${check.token}" }`)
    .join(" ");
  return `${declarations}
${branches} else { "NONE" }`;
}

async function verifySingleOfficeComProgId(
  progId: string,
  activeObjectExpression: string,
): Promise<{ version?: string; activeName?: string } | null> {
  try {
    const result = await executePowerShell(`
          try {
            $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
            "OK|$($app.Version)|$(if (${activeObjectExpression}) { ${activeObjectExpression}.Name } else { '' })"
          } catch { "FAIL" }
        `);
    if (!result.startsWith("OK|")) return null;
    const parts = result.split("|");
    return { version: parts[1] || undefined, activeName: parts[2] || undefined };
  } catch {
    return null;
  }
}
