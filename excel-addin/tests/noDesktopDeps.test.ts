import { readdirSync, readFileSync, statSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

/** Runtime dependency patterns only (not doc mentions of desktop/electron source paths). */
const FORBIDDEN = [
  /from\s+["'][^"']*desktop\//,
  /from\s+["']electron["']/,
  /require\(\s*["']electron["']\s*\)/,
  /require\(\s*["'][^"']*desktop\//,
  /Wengge\.OfficeWorker/,
  /System\.Runtime\.InteropServices/,
  /Microsoft\.Office\.Interop/,
  /node:child_process/,
  /from\s+["']child_process["']/,
  /node-adodb/i,
  /edge-js/i,
  /@dotnet\//,
];

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
  it("source tree does not import desktop runtime or native office bridges", () => {
    const files = walk(root).filter((file) => {
      // Self + build-time package CLI (spawns npm run build only; not runtime host bridge).
      if (file.includes(`${path.sep}tests${path.sep}noDesktopDeps.test.ts`)) return false;
      if (file.endsWith(`${path.sep}scripts${path.sep}package-prod.mjs`)) return false;
      return true;
    });
    const offenders: string[] = [];
    for (const file of files) {
      const text = readFileSync(file, "utf8");
      for (const pattern of FORBIDDEN) {
        if (pattern.test(text)) {
          offenders.push(`${path.relative(root, file)} :: ${pattern}`);
        }
      }
    }
    expect(offenders).toEqual([]);
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
});
