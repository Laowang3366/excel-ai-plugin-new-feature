import path from "path";
import { describe, expect, test } from "vitest";
import { existsSync } from "fs";
import {
  detectScriptEngine,
  executeSmart,
  resetEngineCache,
} from "./scriptEngine";
import {
  getEmbeddedPythonPath,
  getEmbeddedPythonPathCandidates,
  getPythonPath,
} from "./python";

describe("getEmbeddedPythonPathCandidates", () => {
  test("includes the desktop-local embedded Python path for development runs", () => {
    const expected = path.normalize(path.join(process.cwd(), "python", "python.exe"));

    expect(getEmbeddedPythonPathCandidates()).toContain(expected);
  });

  test("uses the desktop-local embedded Python when it is installed", () => {
    const expected = path.normalize(path.join(process.cwd(), "python", "python.exe"));

    if (!existsSync(expected)) return;

    expect(getEmbeddedPythonPath()).toBe(expected);
    expect(getPythonPath()).toBe(expected);
  });

  test("detects Python when the embedded runtime has xlwings installed", async () => {
    const expected = path.normalize(path.join(process.cwd(), "python", "python.exe"));

    if (!existsSync(expected)) return;

    resetEngineCache();
    await expect(detectScriptEngine()).resolves.toBe("python");
    resetEngineCache();
  });
});

describe("executeSmart", () => {
  test("falls back to PowerShell when Python returns no output", async () => {
    if (process.platform !== "win32") return;

    const result = await executeSmart(
      "",
      '"fallback"',
      15000,
    );

    expect(result).toEqual({ result: "fallback", engine: "powershell" });
  });
});
