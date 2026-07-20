/**
 * Pre-mutation validation of a WPS JSA package directory.
 */
import { createHash } from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import {
  assertNoSensitiveDistPaths,
  listFilesRecursiveStrict,
} from "./packageProdCore.mjs";
import { assertNoRuntimeDesktopDepsInPackageFiles } from "./runtimeDesktopDeps.mjs";
import {
  assertRealDirectory,
  assertRealFile,
  assertAncestryReal,
} from "./wpsJsaInstallPaths.mjs";
import {
  normalizeWpsGitSha,
  validateWpsIndexHtml,
  validateWpsSourceBundle,
  WPS_ADDON_DIRECTORY,
  WPS_ADDON_NAME,
  WPS_ENTRY_SCRIPT,
  WPS_PUBLISH_URL,
} from "./wpsJsaPackage.mjs";

const SAFE_VERSION_RE = /^\d+\.\d+\.\d+(?:\.\d+)?$/;

function readRegularFile(filePath) {
  assertRealFile(filePath, "package file");
  return fs.readFileSync(filePath, "utf8");
}

function readRegularBuffer(filePath) {
  assertRealFile(filePath, "package file");
  return fs.readFileSync(filePath);
}

/**
 * Parse SHA256SUMS.txt: "<hex>  <relpath>" (two spaces).
 * @returns {Map<string, string>} rel -> hash
 */
export function parseSha256Sums(text) {
  const map = new Map();
  const lines = String(text).replace(/\r\n/g, "\n").split("\n");
  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i];
    if (line.trim() === "") continue;
    const m = /^([0-9a-f]{64})  (.+)$/.exec(line);
    if (!m) {
      throw new Error(`invalid SHA256SUMS line ${i + 1}: ${line}`);
    }
    const hash = m[1];
    const rel = m[2].replace(/\\/g, "/");
    if (path.isAbsolute(rel) || rel.startsWith("/") || /^[A-Za-z]:/.test(rel)) {
      throw new Error(`absolute path forbidden in SHA256SUMS: ${rel}`);
    }
    if (rel.includes("\\")) {
      throw new Error(`backslash path forbidden in SHA256SUMS: ${rel}`);
    }
    if (rel.split("/").some((s) => s === "" || s === "." || s === "..")) {
      throw new Error(`path traversal forbidden in SHA256SUMS: ${rel}`);
    }
    if (rel === "SHA256SUMS.txt") {
      throw new Error("SHA256SUMS must not list itself");
    }
    if (map.has(rel)) {
      throw new Error(`duplicate SHA256SUMS entry: ${rel}`);
    }
    map.set(rel, hash);
  }
  if (map.size === 0) throw new Error("SHA256SUMS is empty");
  return map;
}

function sha256File(absPath) {
  return createHash("sha256").update(readRegularBuffer(absPath)).digest("hex");
}

/**
 * @param {string} packageDir
 * @returns {{
 *   packageDir: string,
 *   buildInfo: object,
 *   hashes: Map<string, string>,
 *   files: string[],
 *   addonDir: string,
 * }}
 */
export function validateWpsPackageDir(packageDir) {
  const root = assertRealDirectory(path.resolve(packageDir), "packageDir");
  assertAncestryReal(root, root);

  const buildInfoPath = path.join(root, "BUILD_INFO.json");
  const sumsPath = path.join(root, "SHA256SUMS.txt");
  const publishPath = path.join(root, "publish.xml");
  const addonDir = path.join(root, WPS_ADDON_DIRECTORY);

  assertRealFile(buildInfoPath, "BUILD_INFO.json");
  assertRealFile(sumsPath, "SHA256SUMS.txt");
  assertRealFile(publishPath, "publish.xml");
  assertRealDirectory(addonDir, "addonDirectory");

  let buildInfo;
  try {
    buildInfo = JSON.parse(readRegularFile(buildInfoPath));
  } catch {
    throw new Error("BUILD_INFO.json is not valid JSON");
  }
  if (buildInfo?.target !== "wps-jsa") {
    throw new Error(`BUILD_INFO.target must be wps-jsa, got ${buildInfo?.target}`);
  }
  if (buildInfo.addonName !== WPS_ADDON_NAME) {
    throw new Error("BUILD_INFO.addonName mismatch");
  }
  if (buildInfo.addonDirectory !== WPS_ADDON_DIRECTORY) {
    throw new Error("BUILD_INFO.addonDirectory must equal WPS_ADDON_DIRECTORY");
  }
  if (buildInfo.publishUrl !== WPS_PUBLISH_URL) {
    throw new Error("BUILD_INFO.publishUrl mismatch");
  }
  if (!SAFE_VERSION_RE.test(String(buildInfo.packageVersion ?? ""))) {
    throw new Error(`invalid BUILD_INFO.packageVersion: ${buildInfo.packageVersion}`);
  }
  const gitSha = normalizeWpsGitSha(String(buildInfo.gitSha ?? "unknown"));
  buildInfo = { ...buildInfo, gitSha };

  const hashes = parseSha256Sums(readRegularFile(sumsPath));
  const files = listFilesRecursiveStrict(root).sort();
  assertNoSensitiveDistPaths(files);

  const required = [
    "BUILD_INFO.json",
    "publish.xml",
    `${WPS_ADDON_DIRECTORY}/index.html`,
    `${WPS_ADDON_DIRECTORY}/manifest.xml`,
    `${WPS_ADDON_DIRECTORY}/ribbon.xml`,
    `${WPS_ADDON_DIRECTORY}/${WPS_ENTRY_SCRIPT}`,
  ];
  for (const rel of required) {
    if (!files.includes(rel)) throw new Error(`package missing required file: ${rel}`);
    if (rel !== "BUILD_INFO.json" && !hashes.has(rel) && rel !== "SHA256SUMS.txt") {
      // BUILD_INFO is hashed in package sums
    }
  }
  // SHA256SUMS lists all package files except itself
  const hashedExpected = files.filter((f) => f !== "SHA256SUMS.txt");
  for (const rel of hashedExpected) {
    if (!hashes.has(rel)) throw new Error(`SHA256SUMS missing entry: ${rel}`);
  }
  for (const rel of hashes.keys()) {
    if (!hashedExpected.includes(rel)) {
      throw new Error(`SHA256SUMS lists file not present (or extra): ${rel}`);
    }
  }

  for (const [rel, expected] of hashes) {
    const abs = path.join(root, ...rel.split("/"));
    assertRealFile(abs, rel);
    const actual = sha256File(abs);
    if (actual !== expected) {
      throw new Error(`hash mismatch for ${rel}`);
    }
  }

  // Re-validate addon manifests from package (do not trust BUILD_INFO alone)
  const source = {
    sourceDir: addonDir,
    manifestXml: readRegularFile(path.join(addonDir, "manifest.xml")),
    ribbonXml: readRegularFile(path.join(addonDir, "ribbon.xml")),
    entryScript: readRegularFile(path.join(addonDir, WPS_ENTRY_SCRIPT)),
    publishXml: readRegularFile(publishPath),
  };
  const bundleCheck = validateWpsSourceBundle(source);
  if (!bundleCheck.ok) {
    throw new Error(`package source validation failed: ${bundleCheck.errors.join("; ")}`);
  }
  const indexHtml = readRegularFile(path.join(addonDir, "index.html"));
  const indexCheck = validateWpsIndexHtml(indexHtml);
  if (!indexCheck.ok) {
    throw new Error(`package index.html invalid: ${indexCheck.errors.join("; ")}`);
  }

  const textArtifacts = files
    .filter((r) => /\.(js|mjs|cjs|html|css|json|xml|md|txt)$/i.test(r))
    .map((rel) => ({
      relativePath: rel,
      content: readRegularFile(path.join(root, ...rel.split("/"))),
    }));
  assertNoRuntimeDesktopDepsInPackageFiles(textArtifacts);

  return {
    packageDir: root,
    buildInfo,
    hashes,
    files: hashedExpected,
    addonDir,
  };
}

export function hashMapToObject(map) {
  const out = {};
  for (const [k, v] of [...map.entries()].sort((a, b) => (a[0] < b[0] ? -1 : 1))) {
    out[k] = v;
  }
  return out;
}

export function packageDigest(hashes) {
  const rows = [...hashes.entries()]
    .sort((a, b) => (a[0] < b[0] ? -1 : 1))
    .map(([rel, hash]) => `${hash}  ${rel}`);
  return createHash("sha256").update(rows.join("\n") + "\n").digest("hex");
}
