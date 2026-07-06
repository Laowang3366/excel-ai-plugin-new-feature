import { executePowerShell } from "../../../automation/powershell";
import type { SpreadsheetHost } from "./connectionMetadata";

export interface ExcelProcessStatus {
  running: boolean;
  host: SpreadsheetHost | "unknown";
  availableHosts: SpreadsheetHost[];
}

export interface ExcelComAvailability {
  available: boolean;
  version?: string;
  workbookName?: string;
}

export async function detectExcelProcess(): Promise<ExcelProcessStatus> {
  try {
    const result = await executePowerShell(`
      $excel = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
      $wps = Get-Process -Name "et" -ErrorAction SilentlyContinue
      $hosts = @()
      if ($excel) { $hosts += "EXCEL" }
      if ($wps)   { $hosts += "WPS" }
      if ($hosts.Count -eq 0) { "NONE" } else { $hosts -join "," }
    `);
    const trimmed = result.trim();
    if (trimmed === "NONE" || !trimmed) {
      return { running: false, host: "unknown", availableHosts: [] };
    }
    const hosts = trimmed.split(",").map((host: string) =>
      host.trim() === "WPS" ? "wps" : "excel"
    ) as SpreadsheetHost[];

    return {
      running: true,
      host: hosts[0],
      availableHosts: hosts,
    };
  } catch {
    return { running: false, host: "unknown", availableHosts: [] };
  }
}

export async function verifyExcelComAvailable(host: SpreadsheetHost): Promise<ExcelComAvailability> {
  try {
    const progId = host === "wps" ? "Ket.Application" : "Excel.Application";
    const result = await executePowerShell(`
      try {
        $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
        $ver = $app.Version
        $wb = $app.ActiveWorkbook
        $wbName = if ($wb) { $wb.Name } else { '' }
        "OK|$ver|$wbName"
      } catch {
        "FAIL"
      }
    `);
    if (result.startsWith("OK|")) {
      const parts = result.split("|");
      return {
        available: true,
        version: parts[1] || undefined,
        workbookName: parts[2] || undefined,
      };
    }
    return { available: false };
  } catch {
    return { available: false };
  }
}
