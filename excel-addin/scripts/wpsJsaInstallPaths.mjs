/**
 * Safe path resolution for WPS JSA install-time CLI (Node only).
 */
import { randomBytes } from "node:crypto";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { LEGACY_OWN_ADDON_DIRECTORY, WPS_ADDON_DIRECTORY } from "./wpsJsaPackage.mjs";

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

/** lstat if path or dangling symlink exists; null only when truly absent. */
export function lstatIfPresent(absPath) {
  try {
    return fs.lstatSync(absPath);
  } catch (error) {
    if (error && error.code === "ENOENT") return null;
    throw error;
  }
}

export function assertRealDirectory(absPath, label = "path") {
  const resolved = path.resolve(absPath);
  const st = lstatIfPresent(resolved);
  if (!st) throw new Error(`${label} does not exist: ${resolved}`);
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
  const st = lstatIfPresent(resolved);
  if (!st) throw new Error(`${label} does not exist: ${resolved}`);
  if (st.isSymbolicLink()) {
    throw new Error(`${label} must not be a symlink/junction: ${resolved}`);
  }
  if (!st.isFile()) {
    throw new Error(`${label} must be a regular file: ${resolved}`);
  }
  return resolved;
}

export function assertAncestryReal(absPath, stopAt) {
  const target = path.resolve(absPath);
  const rootDir = path.resolve(stopAt);
  if (target !== rootDir && !target.startsWith(rootDir + path.sep)) {
    throw new Error(`path escapes allowed root: ${target}`);
  }
  const rel = path.relative(rootDir, target);
  let current = rootDir;
  if (lstatIfPresent(rootDir)) {
    assertRealDirectory(rootDir, "root");
  }
  if (rel === "" || rel === ".") return target;
  for (const segment of rel.split(path.sep)) {
    current = path.join(current, segment);
    const st = lstatIfPresent(current);
    if (!st) break;
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
  const legacyAddonDir = path.join(jsaddons, LEGACY_OWN_ADDON_DIRECTORY);
  const publishXml = path.join(jsaddons, "publish.xml");
  const stateFile = path.join(jsaddons, STATE_FILE_NAME);
  assertInside(appData, jsaddons, "jsaddons");
  assertInside(jsaddons, addonDir, "addonDir");
  assertInside(jsaddons, legacyAddonDir, "legacyAddonDir");
  assertInside(jsaddons, publishXml, "publish.xml");
  assertInside(jsaddons, stateFile, "state file");
  return { appData, jsaddons, addonDir, legacyAddonDir, publishXml, stateFile };
}

export function ensureJsaddonsDir(jsaddons, appData) {
  assertInside(appData, jsaddons, "jsaddons");
  if (lstatIfPresent(jsaddons)) {
    assertRealDirectory(jsaddons, "jsaddons");
    assertAncestryReal(jsaddons, appData);
    return;
  }
  const segments = path.relative(appData, jsaddons).split(path.sep);
  let current = appData;
  if (!lstatIfPresent(appData)) {
    fs.mkdirSync(appData, { recursive: true });
  }
  assertRealDirectory(appData, "appData");
  for (const segment of segments) {
    current = path.join(current, segment);
    if (lstatIfPresent(current)) {
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

/** Reserve exclusive path for rename target: create empty then unlink, return free path. */
export function reserveExclusivePath(jsaddons, prefix) {
  const full = exclusiveTempFile(jsaddons, prefix);
  fs.unlinkSync(full);
  return full;
}

export function reserveExclusiveDirPath(jsaddons, prefix) {
  const full = exclusiveTempDir(jsaddons, prefix);
  fs.rmdirSync(full);
  return full;
}

export function stagingDir(jsaddons) {
  return exclusiveTempDir(jsaddons, STAGING_PREFIX);
}

export function prevAddonDir(jsaddons) {
  return exclusiveTempDir(jsaddons, PREV_PREFIX);
}

/** Timestamp + random own-prefix backup path (exclusive create). */
export function ownPublishBackupPath(jsaddons) {
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  return exclusiveTempFile(jsaddons, `${PUBLISH_BACKUP_PREFIX}${stamp}.`);
}

/** Read-only validate own-prefix backups (no symlink/non-regular). */
export function assertOwnPublishBackupSurface(jsaddons) {
  assertRealDirectory(jsaddons, "jsaddons");
  for (const name of fs.readdirSync(jsaddons)) {
    if (!name.startsWith(PUBLISH_BACKUP_PREFIX)) continue;
    const full = path.join(jsaddons, name);
    assertInside(jsaddons, full, "backup");
    const st = lstatIfPresent(full);
    if (!st) continue;
    if (st.isSymbolicLink() || !st.isFile()) {
      throw new Error(`own publish backup is not a regular file: ${full}`);
    }
  }
}

/**
 * Rotate only our backup prefix by mtimeMs (tie-break name); never touch publish.xml.bak.*.
 * Best-effort after successful commit — caller should not roll back install on failure.
 */
export function rotateOwnPublishBackups(jsaddons) {
  assertRealDirectory(jsaddons, "jsaddons");
  assertOwnPublishBackupSurface(jsaddons);
  const entries = fs
    .readdirSync(jsaddons)
    .filter((name) => name.startsWith(PUBLISH_BACKUP_PREFIX))
    .map((name) => {
      const full = path.join(jsaddons, name);
      const st = fs.lstatSync(full);
      return { name, full, mtimeMs: st.mtimeMs };
    })
    .sort((a, b) => {
      if (a.mtimeMs !== b.mtimeMs) return a.mtimeMs - b.mtimeMs; // oldest first
      return a.name < b.name ? -1 : a.name > b.name ? 1 : 0;
    });
  while (entries.length > MAX_OWN_PUBLISH_BACKUPS) {
    const oldest = entries.shift();
    assertInside(jsaddons, oldest.full, "backup");
    const st = lstatIfPresent(oldest.full);
    if (!st || st.isSymbolicLink() || !st.isFile()) {
      throw new Error(`refusing to remove non-regular backup: ${oldest.full}`);
    }
    fs.unlinkSync(oldest.full);
  }
}

export function safeRmInside(parent, target) {
  const p = path.resolve(parent);
  const t = path.resolve(target);
  if (t === p) throw new Error("refusing to remove parent directory");
  assertInside(p, t, "remove target");
  const st = lstatIfPresent(t);
  if (!st) return;
  if (st.isSymbolicLink()) {
    throw new Error(`refusing to remove symlink: ${t}`);
  }
  assertAncestryReal(path.dirname(t), p);
  if (t === p) throw new Error("refusing to remove jsaddons root");
  fs.rmSync(t, { recursive: true, force: false });
}

export function safeRenameInside(parent, from, to) {
  assertInside(parent, from, "rename from");
  assertInside(parent, to, "rename to");
  const fromSt = lstatIfPresent(from);
  if (!fromSt) throw new Error(`rename source missing: ${from}`);
  if (fromSt.isSymbolicLink()) {
    throw new Error(`refusing to rename symlink: ${from}`);
  }
  const toSt = lstatIfPresent(to);
  if (toSt) {
    throw new Error(`rename destination already exists: ${to}`);
  }
  fs.renameSync(from, to);
}

/** Preflight existing install surface (detects dangling symlinks). */
export function preflightExistingSurface(layout) {
  const { jsaddons, addonDir, publishXml, stateFile } = layout;
  assertRealDirectory(jsaddons, "jsaddons");
  for (const [p, label, kind] of [
    [publishXml, "publish.xml", "file"],
    [stateFile, "state", "file"],
    [addonDir, "addonDir", "dir"],
  ]) {
    const st = lstatIfPresent(p);
    if (!st) continue;
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

/**
 * Fail closed if any own-prefix entry is symlink/non-regular (before mutations).
 */
export function preflightOwnPrefixedEntries(jsaddons) {
  assertRealDirectory(jsaddons, "jsaddons");
  const prefixes = [STAGING_PREFIX, PREV_PREFIX, TMP_PREFIX, PUBLISH_BACKUP_PREFIX];
  for (const name of fs.readdirSync(jsaddons)) {
    if (!prefixes.some((p) => name.startsWith(p))) continue;
    const full = path.join(jsaddons, name);
    assertInside(jsaddons, full, "own-prefix entry");
    const st = lstatIfPresent(full);
    if (!st) continue;
    if (st.isSymbolicLink()) {
      throw new Error(`own-prefix entry must not be symlink: ${full}`);
    }
    if (name.startsWith(PUBLISH_BACKUP_PREFIX) || name.startsWith(TMP_PREFIX)) {
      if (!st.isFile()) {
        throw new Error(`own-prefix temp/backup must be regular file: ${full}`);
      }
    } else if (!st.isDirectory() && !st.isFile()) {
      throw new Error(`own-prefix entry has unexpected type: ${full}`);
    }
  }
}

/** Active stage/prev/tmp names currently present (for tests/cleanup checks). */
export function listActiveTempNames(jsaddons) {
  if (!lstatIfPresent(jsaddons)) return [];
  return fs
    .readdirSync(jsaddons)
    .filter(
      (n) =>
        n.startsWith(STAGING_PREFIX) ||
        n.startsWith(PREV_PREFIX) ||
        n.startsWith(TMP_PREFIX),
    )
    .sort();
}

export function tmpdir() {
  return os.tmpdir();
}
