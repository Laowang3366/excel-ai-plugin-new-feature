#!/usr/bin/env node
/**
 * Production static package for Excel add-in.
 * Builds with VITE_BASE, writes prod Office manifest + BUILD_INFO + SHA256SUMS into dist/.
 */
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import {
  assertIndexAssetsUnderBase,
  assertLocalAssetFiles,
  assertNoSensitiveDistPaths,
  assertProductionDistClean,
  buildBuildInfo,
  buildSha256Sums,
  formatSpawnFailure,
  listFilesRecursiveStrict,
  makeArtifactName,
  parseCliArgs,
  resolvePackageInputs,
} from "./packageProdCore.mjs";
import {
  renderOfficeManifest,
  validateOfficeManifest,
} from "./officeManifest.mjs";
import { assertNoRuntimeDesktopDepsInPackageFiles } from "./runtimeDesktopDeps.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const defaultRoot = path.resolve(__dirname, "..");

function usage(code = 1) {
  console.error(`Usage:
  npm run package:prod -- --base-url https://example.com/excel-addin [--version 0.1.0.0] [--vite-base /excel-addin/] [--git-sha <sha>]

Writes production static files into dist/ and prints one JSON summary line to stdout.
`);
  process.exit(code);
}

function assertSafeDistPath(rootDir, distDir) {
  const relative = path.relative(rootDir, distDir);
  if (
    relative === "" ||
    relative === ".." ||
    relative.startsWith(`..${path.sep}`) ||
    path.isAbsolute(relative)
  ) {
    throw new Error(`dist directory must be inside project root: ${distDir}`);
  }
  const rootStat = fs.lstatSync(rootDir);
  if (rootStat.isSymbolicLink() || !rootStat.isDirectory()) {
    throw new Error(`project root must be a real directory: ${rootDir}`);
  }
  let current = rootDir;
  for (const segment of relative.split(path.sep)) {
    current = path.join(current, segment);
    let stat;
    try {
      stat = fs.lstatSync(current);
    } catch (error) {
      if (error?.code === "ENOENT") break;
      throw error;
    }
    if (stat.isSymbolicLink()) {
      throw new Error(`dist path must not contain symlinks: ${current}`);
    }
    if (!stat.isDirectory()) {
      throw new Error(`dist path component must be a directory: ${current}`);
    }
  }
}

function removeDist(rootDir, distDir) {
  assertSafeDistPath(rootDir, distDir);
  fs.rmSync(distDir, { recursive: true, force: true });
}

function runNpmBuild(rootDir, viteBase) {
  const isWin = process.platform === "win32";
  const npmExecPath = process.env.npm_execpath;
  const cmd = npmExecPath ? process.execPath : isWin ? "npm.cmd" : "npm";
  const cmdArgs = npmExecPath
    ? [npmExecPath, "run", "build"]
    : ["run", "build"];
  const env = { ...process.env, VITE_BASE: viteBase, VITE_DEV_HTTP: "1" };
  // Avoid HTTPS cert dependency during build (build never needs certs).
  const result = spawnSync(cmd, cmdArgs, {
    cwd: rootDir,
    env,
    encoding: "utf8",
    shell: false,
  });
  if (result.error || result.signal || result.status !== 0) {
    console.error(result.stdout || "");
    console.error(result.stderr || "");
    throw new Error(formatSpawnFailure(result));
  }
}

export function createPackage(args, env = process.env) {
  if (!args?.baseUrl) throw new Error("--base-url is required");
  const rootDir = path.resolve(args.rootDir || defaultRoot);
  const distDir = path.resolve(args.distDir || path.join(rootDir, "dist"));
  assertSafeDistPath(rootDir, distDir);
  const pkg = JSON.parse(fs.readFileSync(path.join(rootDir, "package.json"), "utf8"));
  const rawGitSha =
    args.gitSha ||
    env.GITHUB_SHA ||
    env.GIT_SHA ||
    "unknown";
  const gitSha = /^[0-9a-f]{7,64}$/i.test(String(rawGitSha).trim())
    ? String(rawGitSha).trim()
    : "unknown";

  const resolved = resolvePackageInputs({
    baseUrl: args.baseUrl,
    version: args.version,
    viteBase: args.viteBase,
    packageJsonVersion: pkg.version,
  });

  let packagingStarted = false;
  try {
    if (!args.skipBuild) {
      removeDist(rootDir, distDir);
      packagingStarted = true;
      runNpmBuild(rootDir, resolved.viteBase);
    } else {
      packagingStarted = true;
    }
    if (!fs.existsSync(distDir)) {
      throw new Error(`dist not found: ${distDir}`);
    }

    const indexPath = path.join(distDir, "index.html");
    if (!fs.existsSync(indexPath)) {
      throw new Error("dist/index.html missing after build");
    }
    const indexHtml = fs.readFileSync(indexPath, "utf8");
    const localAssets = assertIndexAssetsUnderBase(indexHtml, resolved.viteBase);
    assertLocalAssetFiles(distDir, ["index.html", ...localAssets]);

    const template = fs.readFileSync(
      path.join(rootDir, "manifest/templates/office-excel-manifest.template.xml"),
      "utf8",
    );
    const manifestXml = renderOfficeManifest({
      mode: "prod",
      baseUrl: resolved.baseUrl,
      version: resolved.version,
      template,
    });
    const manifestValidation = validateOfficeManifest(manifestXml, { mode: "prod" });
    if (!manifestValidation.ok) {
      throw new Error(manifestValidation.errors.join("; "));
    }
    fs.writeFileSync(
      path.join(distDir, "office-excel-manifest.xml"),
      manifestXml,
      "utf8",
    );

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

    assertLocalAssetFiles(distDir, [
      "index.html",
      ...localAssets,
      "assets/icon-16.png",
      "assets/icon-32.png",
      "assets/icon-64.png",
      "assets/icon-80.png",
      "office-excel-manifest.xml",
      "BUILD_INFO.json",
    ]);

    // Hash all package files except SHA256SUMS itself.
    let rels = listFilesRecursiveStrict(distDir);
    assertNoSensitiveDistPaths(rels);
    const files = rels
      .filter((r) => r !== "SHA256SUMS.txt")
      .map((rel) => ({
        relativePath: rel,
        content: fs.readFileSync(path.join(distDir, rel)),
      }));
    const sums = buildSha256Sums(files);
    fs.writeFileSync(path.join(distDir, "SHA256SUMS.txt"), sums, "utf8");

    // Final scan including sums file name (ok) but still no secrets/symlinks.
    rels = listFilesRecursiveStrict(distDir);
    assertNoSensitiveDistPaths(rels);
    assertProductionDistClean({
      distDir,
      baseUrl: resolved.baseUrl,
      viteBase: resolved.viteBase,
      relativePaths: rels,
    });

    const textArtifacts = rels
      .filter((r) => /\.(js|mjs|cjs|html|css|json|xml|md|txt)$/i.test(r))
      .map((rel) => ({
        relativePath: rel,
        content: fs.readFileSync(path.join(distDir, rel), "utf8"),
      }));
    assertNoRuntimeDesktopDepsInPackageFiles(textArtifacts);

    return {
      ok: true,
      artifactName: makeArtifactName(resolved.version, gitSha),
      baseUrl: resolved.baseUrl,
      viteBase: resolved.viteBase,
      version: resolved.version,
      gitSha,
      distDir,
      files: rels.slice().sort(),
    };
  } catch (error) {
    if (packagingStarted) removeDist(rootDir, distDir);
    throw error;
  }
}

function main() {
  let args;
  try {
    args = parseCliArgs(process.argv.slice(2));
  } catch (error) {
    console.error(String(error?.message || error));
    usage(1);
  }
  if (args.help) usage(0);
  try {
    const summary = createPackage(args);
    // Single machine-readable line for CI.
    process.stdout.write(`${JSON.stringify(summary)}\n`);
  } catch (error) {
    console.error(String(error?.message || error));
    process.exitCode = 1;
  }
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main();
}
