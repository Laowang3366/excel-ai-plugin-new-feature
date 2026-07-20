#!/usr/bin/env node
/** Build the shared task pane as a local-file WPS JSA jsaddons package. */
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import {
  assertLocalAssetFiles,
  assertNoSensitiveDistPaths,
  buildSha256Sums,
  formatSpawnFailure,
  listFilesRecursiveStrict,
} from "./packageProdCore.mjs";
import {
  makeWpsArtifactName,
  normalizeWpsGitSha,
  prepareWpsIndexHtml,
  validateWpsIndexHtml,
  validateWpsSourceBundle,
  WPS_ADDON_DIRECTORY,
  WPS_ADDON_NAME,
  WPS_ENTRY_SCRIPT,
  WPS_PUBLISH_URL,
} from "./wpsJsaPackage.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const defaultRoot = path.resolve(__dirname, "..");

function usage(code = 1) {
  console.error(`Usage:
  npm run package:wps -- [--git-sha <sha>]

Writes a local-file WPS jsaddons package into dist/.
`);
  process.exit(code);
}

function parseArgs(argv) {
  const out = { gitSha: null, help: false };
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === "--git-sha") out.gitSha = argv[++i] ?? null;
    else if (arg === "-h" || arg === "--help") out.help = true;
    else throw new Error(`Unknown arg: ${arg}`);
  }
  return out;
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

function runWpsBuild(rootDir) {
  const npmExecPath = process.env.npm_execpath;
  const command = npmExecPath
    ? process.execPath
    : process.platform === "win32"
      ? "npm.cmd"
      : "npm";
  const args = npmExecPath
    ? [npmExecPath, "run", "build:wps"]
    : ["run", "build:wps"];
  const result = spawnSync(command, args, {
    cwd: rootDir,
    env: { ...process.env, VITE_DEV_HTTP: "1" },
    encoding: "utf8",
    shell: false,
  });
  if (result.error || result.signal || result.status !== 0) {
    console.error(result.stdout || "");
    console.error(result.stderr || "");
    throw new Error(formatSpawnFailure(result));
  }
}

function readRegularFile(filePath) {
  const stat = fs.lstatSync(filePath);
  if (stat.isSymbolicLink() || !stat.isFile()) {
    throw new Error(`WPS source must be a regular file: ${filePath}`);
  }
  return fs.readFileSync(filePath, "utf8");
}

function readSourceBundle(rootDir) {
  const sourceDir = path.join(rootDir, "manifest/wps-jsa");
  return {
    sourceDir,
    manifestXml: readRegularFile(path.join(sourceDir, "manifest.xml")),
    ribbonXml: readRegularFile(path.join(sourceDir, "ribbon.xml")),
    entryScript: readRegularFile(path.join(sourceDir, WPS_ENTRY_SCRIPT)),
    publishXml: readRegularFile(path.join(sourceDir, "publish.xml")),
  };
}

function moveTaskPaneIntoAddon(distDir) {
  const entries = fs.readdirSync(distDir, { withFileTypes: true });
  if (entries.some((entry) => entry.name === WPS_ADDON_DIRECTORY)) {
    throw new Error(`Vite output collides with WPS add-on directory: ${WPS_ADDON_DIRECTORY}`);
  }
  for (const entry of entries) {
    if (entry.isSymbolicLink() || (!entry.isFile() && !entry.isDirectory())) {
      throw new Error(`unsupported Vite output entry: ${entry.name}`);
    }
  }
  const addonDir = path.join(distDir, WPS_ADDON_DIRECTORY);
  fs.mkdirSync(addonDir);
  for (const entry of entries) {
    fs.renameSync(path.join(distDir, entry.name), path.join(addonDir, entry.name));
  }
  return addonDir;
}

function writePackageMetadata(distDir, pkg, gitSha) {
  const info = {
    target: "wps-jsa",
    packageVersion: String(pkg.version),
    gitSha,
    addonName: WPS_ADDON_NAME,
    addonDirectory: WPS_ADDON_DIRECTORY,
    publishUrl: WPS_PUBLISH_URL,
    assetBase: "./",
  };
  fs.writeFileSync(
    path.join(distDir, "BUILD_INFO.json"),
    `${JSON.stringify(info, null, 2)}\n`,
    "utf8",
  );
}

export function createWpsPackage(args = {}, env = process.env) {
  const rootDir = path.resolve(args.rootDir || defaultRoot);
  const distDir = path.resolve(args.distDir || path.join(rootDir, "dist"));
  assertSafeDistPath(rootDir, distDir);

  const pkg = JSON.parse(readRegularFile(path.join(rootDir, "package.json")));
  const rawGitSha = args.gitSha || env.GITHUB_SHA || env.GIT_SHA || "unknown";
  const gitSha = normalizeWpsGitSha(rawGitSha);
  const source = readSourceBundle(rootDir);
  const sourceValidation = validateWpsSourceBundle(source);
  if (!sourceValidation.ok) {
    throw new Error(sourceValidation.errors.join("; "));
  }

  let packagingStarted = false;
  try {
    if (!args.skipBuild) {
      removeDist(rootDir, distDir);
      packagingStarted = true;
      runWpsBuild(rootDir);
    } else {
      packagingStarted = true;
    }
    if (!fs.existsSync(distDir)) throw new Error(`dist not found: ${distDir}`);

    const addonDir = moveTaskPaneIntoAddon(distDir);
    const indexPath = path.join(addonDir, "index.html");
    if (!fs.existsSync(indexPath)) throw new Error("Vite output is missing index.html");
    const preparedIndex = prepareWpsIndexHtml(readRegularFile(indexPath));
    const indexValidation = validateWpsIndexHtml(preparedIndex);
    if (!indexValidation.ok) throw new Error(indexValidation.errors.join("; "));
    fs.writeFileSync(indexPath, preparedIndex, "utf8");

    fs.copyFileSync(path.join(source.sourceDir, "manifest.xml"), path.join(addonDir, "manifest.xml"));
    fs.copyFileSync(path.join(source.sourceDir, "ribbon.xml"), path.join(addonDir, "ribbon.xml"));
    fs.copyFileSync(
      path.join(source.sourceDir, WPS_ENTRY_SCRIPT),
      path.join(addonDir, WPS_ENTRY_SCRIPT),
    );
    fs.copyFileSync(path.join(source.sourceDir, "publish.xml"), path.join(distDir, "publish.xml"));
    writePackageMetadata(distDir, pkg, gitSha);

    assertLocalAssetFiles(addonDir, [
      "index.html",
      ...indexValidation.assets,
      "assets/icon-16.png",
      "assets/icon-32.png",
      "assets/icon-64.png",
      "assets/icon-80.png",
      "manifest.xml",
      "ribbon.xml",
      WPS_ENTRY_SCRIPT,
    ]);

    let files = listFilesRecursiveStrict(distDir);
    assertNoSensitiveDistPaths(files);
    const sums = buildSha256Sums(
      files
        .filter((relativePath) => relativePath !== "SHA256SUMS.txt")
        .map((relativePath) => ({
          relativePath,
          content: fs.readFileSync(path.join(distDir, relativePath)),
        })),
    );
    fs.writeFileSync(path.join(distDir, "SHA256SUMS.txt"), sums, "utf8");
    files = listFilesRecursiveStrict(distDir).sort();
    assertNoSensitiveDistPaths(files);

    return {
      ok: true,
      artifactName: makeWpsArtifactName(String(pkg.version), gitSha),
      version: String(pkg.version),
      gitSha,
      distDir,
      addonDirectory: WPS_ADDON_DIRECTORY,
      files,
    };
  } catch (error) {
    if (packagingStarted) removeDist(rootDir, distDir);
    throw error;
  }
}

function main() {
  let args;
  try {
    args = parseArgs(process.argv.slice(2));
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    usage(1);
  }
  if (args.help) usage(0);
  try {
    const summary = createWpsPackage(args);
    process.stdout.write(`${JSON.stringify(summary)}\n`);
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    process.exitCode = 1;
  }
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main();
}
