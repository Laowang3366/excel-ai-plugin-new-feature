/**
 * Safe path resolution for WPS JSA install-time CLI (Node only).
 */
import { randomBytes } from "node:crypto";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { WPS_ADDON_DIRECTORY } from "./wpsJsaPackage.mjs";

export const JSADDONS_REL = path.join("kingsoft", "wps", "jsaddons");
export const STATE_FILE_NAME = "wengge-excel-ai-addin-install-state.json";
export const PUBLISH_BACKUP_PREFIX = "publish.xml.wengge-excel-ai.bak.";
export const STAGING_PREFIX = ".wengge-excel-ai-stage-";
export const PREV_PREFIX = ".wengge-excel-ai-prev-";
export const TMP_PREFIX = ".wengge-excel-ai-tmp-";
export const MAX_OWN_PUBLISH_BACKUPS = 10;

export function isWindowsPlatform(platform = process.platform) {
  return platform === "win32";
}

/**
 * @param {{ appData?: string|null, platform?: string, env?: NodeJS.ProcessEnv }} opts
 */
export function resolveAppDataRoot(opts = {}) {
  const platform = opts.platform || process.platform;
  if (opts.appData != null && String(opts.appData).trim() !== "") {
    return path.resolve(String(opts.appData));
  }
  if (!isWindowsPlatform(platform)) {
    throw new Error(
      "Non-Windows hosts require explicit --app-data (refusing default AppData path)",
    );
  }
  const env = opts.env || process.env;
  const appData = env.APPDATA;
  if (!appData || String(appData).trim() === "") {
    throw new Error("APPDATA is not set; pass --app-data explicitly");
  }
  return path.resolve(String(appData));
}

export function assertRealDirectory(absPath, label = "path") {
  const resolved = path.resolve(absPath);
  if (!fs.existsSync(resolved)) {
    throw new Error(`${label} does not exist: ${resolved}`);
  }
  const st = fs.lstatSync(resolved);
  if (st.isSymbolicLink()) {
    throw new Error(`${label} must not be a symlink/junction: ${resolved}`);
  }
  if (!st.isDirectory()) {
    throw new Error(`${label} must be a real directory: ${resolved}`);
  }
  return resolved;
}

export function assertRealFile(absPath, label = "path") {
  const resolved = path.resolve(absPath);
  if (!fs.existsSync(resolved)) {
    throw new Error(`${label} does not exist: ${resolved}`);
  }
  const st = fs.lstatSync(resolved);
  if (st.isSymbolicLink()) {
    throw new Error(`${label} must not be a symlink/junction: ${resolved}`);
  }
  if (!st.isFile()) {
    throw new Error(`${label} must be a regular file: ${resolved}`);
  }
  return resolved;
}

/** Every existing path component under stopAt must be real (no symlink). */
export function assertAncestryReal(absPath, stopAt) {
  const target = path.resolve(absPath);
  const root = path.resolve(stopAt);
  if (target !== root && !target.startsWith(root + path.sep)) {
    throw new Error(`path escapes allowed root: ${target}`);
  }
  const rel = path.relative(root, target);
  let current = root;
  if (fs.existsSync(root)) {
    assertRealDirectory(root, "root");
  }
  if (rel === "" || rel === ".") return target;
  for (const segment of rel.split(path.sep)) {
    current = path.join(current, segment);
    if (!fs.existsSync(current)) break;
    const st = fs.lstatSync(current);
    if (st.isSymbolicLink()) {
      throw new Error(`path must not contain symlink/junction: ${current}`);
    }
    if (st.isDirectory()) continue;
    if (st.isFile() && current === target) continue;
    throw new Error(`unexpected non-directory path component: ${current}`);
  }
  return target;
}

export function assertInside(parent, child, label = "path") {
  const p = path.resolve(parent);
  const c = path.resolve(child);
  if (c !== p && !c.startsWith(p + path.sep)) {
    throw new Error(`${label} escapes parent: ${c}`);
  }
  return c;
}

export function resolveJsaddonsLayout(appDataRoot) {
  const appData = path.resolve(appDataRoot);
  const jsaddons = path.join(appData, JSADDONS_REL);
  const addonDir = path.join(jsaddons, WPS_ADDON_DIRECTORY);
  const publishXml = path.join(jsaddons, "publish.xml");
  const stateFile = path.join(jsaddons, STATE_FILE_NAME);
  assertInside(appData, jsaddons, "jsaddons");
  assertInside(jsaddons, addonDir, "addonDir");
  assertInside(jsaddons, publishXml, "publish.xml");
  assertInside(jsaddons, stateFile, "state file");
  return { appData, jsaddons, addonDir, publishXml, stateFile };
}

export function ensureJsaddonsDir(jsaddons, appData) {
  assertInside(appData, jsaddons, "jsaddons");
  if (fs.existsSync(jsaddons)) {
    assertRealDirectory(jsaddons, "jsaddons");
    assertAncestryReal(jsaddons, appData);
    return;
  }
  const segments = path.relative(appData, jsaddons).split(path.sep);
  let current = appData;
  if (!fs.existsSync(appData)) {
    fs.mkdirSync(appData, { recursive: true });
  }
  assertRealDirectory(appData, "appData");
  for (const segment of segments) {
    current = path.join(current, segment);
    if (fs.existsSync(current)) {
      assertRealDirectory(current, "jsaddons ancestor");
    } else {
      fs.mkdirSync(current);
      assertRealDirectory(current, "jsaddons ancestor");
    }
  }
}

export function randomToken() {
  return randomBytes(12).toString("hex");
}

/** Exclusive create of empty regular file inside jsaddons; returns absolute path. */
export function exclusiveTempFile(jsaddons, prefix = TMP_PREFIX) {
  assertRealDirectory(jsaddons, "jsaddons");
  for (let attempt = 0; attempt < 32; attempt += 1) {
    const full = path.join(jsaddons, `${prefix}${randomToken()}`);
    assertInside(jsaddons, full, "temp");
    try {
      const fd = fs.openSync(full, "wx", 0o600);
      fs.closeSync(fd);
      return full;
    } catch (error) {
      if (error && error.code === "EEXIST") continue;
      throw error;
    }
  }
  throw new Error("unable to allocate exclusive temp file in jsaddons");
}

export function exclusiveTempDir(jsaddons, prefix = STAGING_PREFIX) {
  assertRealDirectory(jsaddons, "jsaddons");
  for (let attempt = 0; attempt < 32; attempt += 1) {
    const full = path.join(jsaddons, `${prefix}${randomToken()}`);
    assertInside(jsaddons, full, "temp dir");
    try {
      fs.mkdirSync(full);
      return full;
    } catch (error) {
      if (error && error.code === "EEXIST") continue;
      throw error;
    }
  }
  throw new Error("unable to allocate exclusive temp directory in jsaddons");
}

export function stagingDir(jsaddons) {
  return exclusiveTempDir(jsaddons, STAGING_PREFIX);
}

export function prevAddonDir(jsaddons) {
  return exclusiveTempDir(jsaddons, PREV_PREFIX);
}

export function ownPublishBackupPath(jsaddons) {
  return exclusiveTempFile(jsaddons, PUBLISH_BACKUP_PREFIX);
}

/** Rotate only our backup prefix; never touch publish.xml.bak.* */
export function rotateOwnPublishBackups(jsaddons) {
  assertRealDirectory(jsaddons, "jsaddons");
  const entries = fs
    .readdirSync(jsaddons, { withFileTypes: true })
    .filter((e) => e.name.startsWith(PUBLISH_BACKUP_PREFIX))
    .map((e) => e.name)
    .sort();
  // fail closed on non-regular/symlink own-prefix entries
  for (const name of entries) {
    const full = path.join(jsaddons, name);
    assertInside(jsaddons, full, "backup");
    const st = fs.lstatSync(full);
    if (st.isSymbolicLink() || !st.isFile()) {
      throw new Error(`own publish backup is not a regular file: ${full}`);
    }
  }
  const files = entries.filter((name) => {
    const full = path.join(jsaddons, name);
    return fs.lstatSync(full).isFile() && !fs.lstatSync(full).isSymbolicLink();
  });
  while (files.length > MAX_OWN_PUBLISH_BACKUPS) {
    const name = files.shift();
    const full = path.join(jsaddons, name);
    assertInside(jsaddons, full, "backup");
    const st = fs.lstatSync(full);
    if (st.isSymbolicLink() || !st.isFile()) {
      throw new Error(`refusing to remove non-regular backup: ${full}`);
    }
    fs.unlinkSync(full);
  }
}

/**
 * Safe recursive remove only if path is strictly inside parent and not parent itself.
 */
export function safeRmInside(parent, target) {
  const p = path.resolve(parent);
  const t = path.resolve(target);
  if (t === p) throw new Error("refusing to remove parent directory");
  assertInside(p, t, "remove target");
  if (!fs.existsSync(t)) return;
  const st = fs.lstatSync(t);
  if (st.isSymbolicLink()) {
    throw new Error(`refusing to remove symlink: ${t}`);
  }
  assertAncestryReal(path.dirname(t), p);
  // re-check not jsaddons root
  if (t === p) throw new Error("refusing to remove jsaddons root");
  fs.rmSync(t, { recursive: true, force: false });
}

export function safeRenameInside(parent, from, to) {
  assertInside(parent, from, "rename from");
  assertInside(parent, to, "rename to");
  if (!fs.existsSync(from)) {
    throw new Error(`rename source missing: ${from}`);
  }
  const st = fs.lstatSync(from);
  if (st.isSymbolicLink()) {
    throw new Error(`refusing to rename symlink: ${from}`);
  }
  if (fs.existsSync(to)) {
    throw new Error(`rename destination already exists: ${to}`);
  }
  fs.renameSync(from, to);
}

/** Preflight existing install surface before mutation. */
export function preflightExistingSurface(layout) {
  const { jsaddons, addonDir, publishXml, stateFile } = layout;
  assertRealDirectory(jsaddons, "jsaddons");
  for (const [p, label, kind] of [
    [publishXml, "publish.xml", "file"],
    [stateFile, "state", "file"],
    [addonDir, "addonDir", "dir"],
  ]) {
    if (!fs.existsSync(p)) continue;
    const st = fs.lstatSync(p);
    if (st.isSymbolicLink()) {
      throw new Error(`${label} must not be a symlink/junction: ${p}`);
    }
    if (kind === "file" && !st.isFile()) {
      throw new Error(`${label} must be a regular file: ${p}`);
    }
    if (kind === "dir" && !st.isDirectory()) {
      throw new Error(`${label} must be a real directory: ${p}`);
    }
  }
}

export function tmpdir() {
  return os.tmpdir();
}
