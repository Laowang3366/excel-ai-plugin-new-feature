import { executePowerShell, psVar } from "../../../automation/powershell";

export interface AcquireOfficeAppScriptOptions {
  progIds: readonly string[];
  allowCreate?: boolean;
  preferredProgId?: string | null;
  missingMessage: string;
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

export function buildAcquireOfficeAppScript(options: AcquireOfficeAppScriptOptions): string {
  const {
    progIds,
    allowCreate = true,
    preferredProgId,
    missingMessage,
  } = options;
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
  return `
$progIds = ${progIdsLiteral(progIds)}
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
