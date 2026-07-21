import { createHash } from "node:crypto";
import { copyFileSync, mkdirSync, readdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, "..");
const repoRoot = path.resolve(root, "..");
const sourceRoot = path.join(
  repoRoot,
  "desktop",
  "electron",
  "agent",
  "prompts",
  "templates",
);
const outDir = path.join(root, "shared", "prompts", "generated");
const templatesDir = path.join(root, "shared", "prompts", "templates");
const manifestPath = path.join(root, "shared", "prompts", "manifest.json");

/** Excel-related prompt sources only (no OCR/Word/PPT-only extras). */
const SOURCES = [
  "system/base.zh-CN.md",
  "system/security.zh-CN.md",
  "scenarios/formula.zh-CN.md",
  "scenarios/general-office.zh-CN.md",
  "scenarios/office-tools.zh-CN.md",
  "scenarios/macro.zh-CN.md",
  "scenarios/ocr-invoice.zh-CN.md",
  "runtime/environment.zh-CN.md",
  "runtime/dynamic-array-enabled.zh-CN.md",
  "runtime/dynamic-array-disabled.zh-CN.md",
];

function sha256(content) {
  return createHash("sha256").update(content).digest("hex");
}

function ensureDir(dir) {
  mkdirSync(dir, { recursive: true });
}

function main() {
  ensureDir(outDir);
  const entries = [];

  for (const relative of SOURCES) {
    const sourcePath = path.join(sourceRoot, relative);
    const targetPath = path.join(outDir, relative);
    ensureDir(path.dirname(targetPath));
    const content = readFileSync(sourcePath);
    copyFileSync(sourcePath, targetPath);
    const sourceRel = path
      .relative(repoRoot, sourcePath)
      .split(path.sep)
      .join("/");
    const overlayPath = path.join(templatesDir, relative);
    let mode = "desktop-identical";
    try {
      statSync(overlayPath);
      mode = "addin-adapted";
    } catch {
      // no overlay
    }
    entries.push({
      id: relative.replace(/\\/g, "/"),
      sourcePath: sourceRel,
      generatedPath: `shared/prompts/generated/${relative.replace(/\\/g, "/")}`,
      sha256: sha256(content),
      bytes: content.byteLength,
      mode,
    });
  }

  const manifest = {
    generatedAt: new Date().toISOString(),
    sourceRoot: "desktop/electron/agent/prompts/templates",
    note: "generated/ keeps desktop-synced copies for audit; templates/ overlays adapt Excel-applicable rules for Office.js/WPS JSA add-in. ocr-invoice synced for audit; add-in uses templates/ overlay (no ocr.parseDocument). Word/PPT remain excluded.",
    files: entries,
  };

  writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`, "utf8");
  console.log(`Synced ${entries.length} prompt files -> ${path.relative(root, outDir)}`);
  console.log(`Manifest -> ${path.relative(root, manifestPath)}`);
}

main();
