import { execFile, spawn } from "child_process";
import { existsSync } from "fs";
import * as path from "path";
import { promisify } from "util";

import type { LaunchOfficeApplicationInput } from "../shared/ipcSchemas";

const execFileAsync = promisify(execFile);

const EXECUTABLE_NAMES: Record<LaunchOfficeApplicationInput, string> = {
  wps: "et.exe",
  excel: "excel.exe",
  word: "winword.exe",
  powerpoint: "powerpnt.exe",
};

export interface OfficeLaunchResult {
  success: boolean;
  error?: string;
}

export function parseRegistryDefaultValue(output: string): string | undefined {
  const match = output.match(/REG_(?:EXPAND_)?SZ\s+(.+?)\s*$/im);
  return match?.[1]?.trim() || undefined;
}

export function getOfficeExecutableName(application: LaunchOfficeApplicationInput): string {
  return EXECUTABLE_NAMES[application];
}

export async function launchOfficeApplication(
  application: LaunchOfficeApplicationInput,
): Promise<OfficeLaunchResult> {
  if (process.platform !== "win32") {
    return { success: false, error: "当前系统不支持启动 Office 程序" };
  }

  const executableName = getOfficeExecutableName(application);
  const executablePath = await resolveExecutablePath(executableName);
  if (!executablePath) {
    return { success: false, error: `未找到 ${getApplicationLabel(application)} 的安装程序` };
  }

  try {
    await spawnDetached(executablePath);
    return { success: true };
  } catch (error) {
    return {
      success: false,
      error: error instanceof Error ? error.message : "启动程序失败",
    };
  }
}

async function resolveExecutablePath(executableName: string): Promise<string | undefined> {
  for (const registryPath of getAppPathRegistryKeys(executableName)) {
    try {
      const { stdout } = await execFileAsync("reg.exe", ["query", registryPath, "/ve"], {
        encoding: "utf8",
        windowsHide: true,
      });
      const registeredPath = parseRegistryDefaultValue(stdout);
      if (registeredPath) {
        const requestedExecutable = path.join(path.dirname(registeredPath), executableName);
        if (existsSync(requestedExecutable)) return requestedExecutable;
        if (existsSync(registeredPath)) return registeredPath;
      }
    } catch {
      // Continue through the remaining registry hives and known install locations.
    }
  }

  return getCommonInstallCandidates(executableName).find(existsSync);
}

function getAppPathRegistryKeys(executableName: string): string[] {
  const suffix = `Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\${executableName}`;
  return [
    `HKCU\\${suffix}`,
    `HKLM\\${suffix}`,
    `HKLM\\Software\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\App Paths\\${executableName}`,
  ];
}

function getCommonInstallCandidates(executableName: string): string[] {
  const candidates: string[] = [];
  for (const programFiles of [process.env.ProgramFiles, process.env["ProgramFiles(x86)"]]) {
    if (!programFiles) continue;
    candidates.push(
      path.join(programFiles, "Microsoft Office", "root", "Office16", executableName),
      path.join(programFiles, "Microsoft Office", "Office16", executableName),
      path.join(programFiles, "Kingsoft", "WPS Office", "office6", executableName),
    );
  }
  return candidates;
}

function spawnDetached(executablePath: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const child = spawn(executablePath, [], {
      detached: true,
      stdio: "ignore",
      windowsHide: false,
    });
    child.once("error", reject);
    child.once("spawn", () => {
      child.unref();
      resolve();
    });
  });
}

function getApplicationLabel(application: LaunchOfficeApplicationInput): string {
  if (application === "wps") return "WPS";
  if (application === "powerpoint") return "Microsoft PowerPoint";
  return `Microsoft ${application === "excel" ? "Excel" : "Word"}`;
}
