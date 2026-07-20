import { readdirSync, readFileSync, statSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import {
  collectRuntimeDesktopDepOffenders,
  findRuntimeDesktopDepHits,
  isDocumentationOnlyDesktopMention,
} from "../scripts/runtimeDesktopDeps.mjs";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

function walk(dir: string, acc: string[] = []): string[] {
  for (const name of readdirSync(dir)) {
    if (
      name === "node_modules" ||
      name === "dist" ||
      name === "generated" ||
      name === "package-lock.json"
    ) {
      continue;
    }
    const full = path.join(dir, name);
    const st = statSync(full);
    if (st.isDirectory()) walk(full, acc);
    else if (/\.(ts|tsx|js|mjs|css|html)$/.test(name)) acc.push(full);
  }
  return acc;
}

describe("no COM/.NET/Electron runtime deps", () => {
  it("distinguishes documentation/prompt text from runtime imports", () => {
    const promptish = [
      'sourcePath: "desktop/electron/agent/prompts/templates/system/base.zh-CN.md"',
      "运行时隔离：COM / .NET Worker / Electron IPC / child_process **禁止且 unsupported**。",
      "desktop/public/wps-jsa-bridge uses Application; not Electron.",
    ].join("\n");
    expect(findRuntimeDesktopDepHits(promptish)).toEqual([]);
    expect(isDocumentationOnlyDesktopMention(promptish)).toBe(true);

    const runtimeSnippets = [
      'import x from "electron";',
      'const cp = require("child_process");',
      'import { spawn } from "node:child_process";',
      'require("edge-js")',
      'from "desktop/electron/foo"',
      "Wengge.OfficeWorker.Office",
      "System.Runtime.InteropServices.Marshal",
    ];
    for (const snippet of runtimeSnippets) {
      expect(findRuntimeDesktopDepHits(snippet).length, snippet).toBeGreaterThan(0);
    }
  });

  it("source tree does not import desktop runtime or native office bridges", () => {
    const files = walk(root).filter((file) => {
      // Self + build-time package CLIs (spawn npm run build* only; not runtime host bridges).
      if (file.includes(`${path.sep}tests${path.sep}noDesktopDeps.test.ts`)) return false;
      if (file.endsWith(`${path.sep}scripts${path.sep}package-prod.mjs`)) return false;
      if (file.endsWith(`${path.sep}scripts${path.sep}package-wps-jsa.mjs`)) return false;
      if (file.endsWith(`${path.sep}scripts${path.sep}runtimeDesktopDeps.mjs`)) return false;
      return true;
    });
    const textFiles = files.map((file) => ({
      relativePath: path.relative(root, file),
      content: readFileSync(file, "utf8"),
    }));
    expect(collectRuntimeDesktopDepOffenders(textFiles)).toEqual([]);
  });

  it("package.json has no electron/dotnet dependencies", () => {
    const pkg = JSON.parse(readFileSync(path.join(root, "package.json"), "utf8")) as {
      dependencies?: Record<string, string>;
      devDependencies?: Record<string, string>;
    };
    const all = { ...pkg.dependencies, ...pkg.devDependencies };
    for (const name of Object.keys(all)) {
      expect(name.toLowerCase()).not.toContain("electron");
      expect(name.toLowerCase()).not.toContain("edge-js");
    }
  });

  it("flags synthetic package artifacts with runtime requires but allows prompt provenance text", () => {
    const clean = collectRuntimeDesktopDepOffenders([
      {
        relativePath: "assets/index.js",
        content:
          'const m={files:[{sourcePath:"desktop/electron/agent/prompts/templates/x.md"}]};\n' +
          '"COM / .NET / Electron / child_process forbidden in runtime";\n',
      },
    ]);
    expect(clean).toEqual([]);

    const dirty = collectRuntimeDesktopDepOffenders([
      {
        relativePath: "assets/evil.js",
        content: 'const {spawn}=require("child_process");\n',
      },
    ]);
    expect(dirty.some((o) => o.includes("require-child_process"))).toBe(true);
  });
});
