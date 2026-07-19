#!/usr/bin/env node
/**
 * Production static package for Excel add-in.
 * Builds with VITE_BASE, writes prod Office manifest + BUILD_INFO + SHA256SUMS into dist/.
 */
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import {
  assertIndexAssetsUnderBase,
  assertNoSensitiveDistPaths,
  buildBuildInfo,
  buildSha256Sums,
  makeArtifactName,
  parseCliArgs,
  resolvePackageInputs,
} from "./packageProdCore.mjs";
import {
  renderOfficeManifest,
  validateOfficeManifest,
} from "./officeManifest.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const defaultRoot = path.resolve(__dirname, "..");

function usage(code = 1) {
  console.error(`Usage:
  npm run package:prod -- --base-url https://example.com/excel-addin [--version 0.1.0.0] [--vite-base /excel-addin/] [--git-sha <sha>]

Writes production static files into dist/ and prints one JSON summary line to stdout.
`);
  process.exit(code);
}

function listFilesRecursive(dir, base = dir) {
  const out = [];
  for (const ent of fs.readdirSync(dir, { withFileTypes: true })) {
    const abs = path.join(dir, ent.name);
    const rel = path.relative(base, abs).replace(/\\/g, "/");
    if (ent.isDirectory()) out.push(...listFilesRecursive(abs, base));
    else out.push(rel);
  }
  return out;
}

function runNpmBuild(rootDir, viteBase) {
  const isWin = process.platform === "win32";
  const cmd = isWin ? "npm.cmd" : "npm";
  const env = { ...process.env, VITE_BASE: viteBase, VITE_DEV_HTTP: "1" };
  // Avoid HTTPS cert dependency during build (build never needs certs).
  const result = spawnSync(cmd, ["run", "build"], {
    cwd: rootDir,
    env,
    encoding: "utf8",
    shell: isWin,
  });
  if (result.status !== 0) {
    console.error(result.stdout || "");
    console.error(result.stderr || "");
    throw new Error(`npm run build failed with status ${result.status}`);
  }
}

function main() {
  let args;
  try {
    args = parseCliArgs(process.argv.slice(2));
  } catch (err) {
    console.error(String(err?.message || err));
    usage(1);
  }
  if (args.help) usage(0);
  if (!args.baseUrl) {
    console.error("--base-url is required");
    usage(1);
  }

  const rootDir = path.resolve(args.rootDir || defaultRoot);
  const distDir = path.resolve(args.distDir || path.join(rootDir, "dist"));
  const pkg = JSON.parse(fs.readFileSync(path.join(rootDir, "package.json"), "utf8"));
  const gitSha =
    args.gitSha ||
    process.env.GITHUB_SHA ||
    process.env.GIT_SHA ||
    "unknown";

  let resolved;
  try {
    resolved = resolvePackageInputs({
      baseUrl: args.baseUrl,
      version: args.version,
      viteBase: args.viteBase,
      packageJsonVersion: pkg.version,
    });
  } catch (err) {
    console.error(String(err?.message || err));
    process.exit(1);
  }

  if (!args.skipBuild) {
    runNpmBuild(rootDir, resolved.viteBase);
  }
  if (!fs.existsSync(distDir)) {
    console.error(`dist not found: ${distDir}`);
    process.exit(1);
  }

  const indexPath = path.join(distDir, "index.html");
  if (!fs.existsSync(indexPath)) {
    console.error("dist/index.html missing after build");
    process.exit(1);
  }
  const indexHtml = fs.readFileSync(indexPath, "utf8");
  try {
    assertIndexAssetsUnderBase(indexHtml, resolved.viteBase);
  } catch (err) {
    console.error(String(err?.message || err));
    process.exit(1);
  }

  const template = fs.readFileSync(
    path.join(rootDir, "manifest/templates/office-excel-manifest.template.xml"),
    "utf8",
  );
  let manifestXml;
  try {
    manifestXml = renderOfficeManifest({
      mode: "prod",
      baseUrl: resolved.baseUrl,
      version: resolved.version,
      template,
    });
    const v = validateOfficeManifest(manifestXml, { mode: "prod" });
    if (!v.ok) throw new Error(v.errors.join("; "));
  } catch (err) {
    console.error(String(err?.message || err));
    process.exit(1);
  }
  fs.writeFileSync(path.join(distDir, "office-excel-manifest.xml"), manifestXml, "utf8");

  const buildInfo = buildBuildInfo({
    gitSha,
    packageVersion: resolved.packageJsonVersion,
    manifestVersion: resolved.version,
    baseUrl: resolved.baseUrl,
    viteBase: resolved.viteBase,
  });
  fs.writeFileSync(
    path.join(distDir, "BUILD_INFO.json"),
    `${JSON.stringify(buildInfo, null, 2)}\n`,
    "utf8",
  );

  // Hash all package files except SHA256SUMS itself.
  let rels = listFilesRecursive(distDir);
  try {
    assertNoSensitiveDistPaths(rels);
  } catch (err) {
    console.error(String(err?.message || err));
    process.exit(1);
  }
  const files = rels
    .filter((r) => r !== "SHA256SUMS.txt")
    .map((rel) => ({
      relativePath: rel,
      content: fs.readFileSync(path.join(distDir, rel)),
    }));
  const sums = buildSha256Sums(files);
  fs.writeFileSync(path.join(distDir, "SHA256SUMS.txt"), sums, "utf8");

  // Final scan including sums file name (ok) but still no secrets.
  rels = listFilesRecursive(distDir);
  try {
    assertNoSensitiveDistPaths(rels);
  } catch (err) {
    console.error(String(err?.message || err));
    process.exit(1);
  }

  const artifactName = makeArtifactName(resolved.version, gitSha);
  const summary = {
    ok: true,
    artifactName,
    baseUrl: resolved.baseUrl,
    viteBase: resolved.viteBase,
    version: resolved.version,
    gitSha,
    distDir,
    files: rels.slice().sort(),
  };
  // Single machine-readable line for CI.
  process.stdout.write(`${JSON.stringify(summary)}\n`);
}

main();
