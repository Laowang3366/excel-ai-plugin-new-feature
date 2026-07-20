/**
 * Path safety for WPS JSA install-time CLI (not task-pane runtime).
 * Fail closed on symlinks/junctions/reparse points and path escape.
 */
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { WPS_ADDON_DIRECTORY } from "./wpsJsaPackage.mjs";

export const JSADDONS_REL = path.join("kingsoft", "wps", "jsaddons");
export const STATE_FILE_NAME = "wengge-excel-ai-addin-install-state.json";
export const PUBLISH_BACKUP_PREFIX = "publish.xml.wengge-excel-ai.bak.";
export const STAGING_PREFIX = ".wengge-excel-ai-stage-";
export const PREV_PREFIX = ".wengge-excel-ai-prev-";
export const MAX_OWN_PUBLISH_BACKUPS = 10;

export function isWindowsPlatform(platform = process.platform) {
  return platform === "win32";
}

/**
 * Resolve AppData root. Non-Windows requires explicit appData.
 * @param {{ appData?: string|null, platform?: string, env?: NodeJS.ProcessEnv }} opts
 */
export function resolveAppDataRoot(opts = {}) {
  const platform = opts.platform ?? process.platform;
  const env = opts.env ?? process.env;
  if (opts.appData != null && String(opts.appData).trim() !== "") {
    return path.resolve(String(opts.appData));
  }
  if (!isWindowsPlatform(platform)) {
    throw new Error(
      "wps install/status/uninstall requires --app-data on non-Windows hosts (refusing default user AppData)",
    );
  }
  const appData = env.APPDATA;
  if (!appData || String(appData).trim() === "") {
    throw new Error("APPDATA is not set; cannot resolve WPS jsaddons location");
  }
  return path.resolve(String(appData));
}

export function assertRealDirectory(absPath, label = "path") {
  const resolved = path.resolve(absPath);
  let st;
  try {
    st = fs.lstatSync(resolved);
  } catch (error) {
    if (error?.code === "ENOENT") {
      throw new Error(`${label} does not exist: ${resolved}`);
    }
    throw error;
  }
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
  let st;
  try {
    st = fs.lstatSync(resolved);
  } catch (error) {
    if (error?.code === "ENOENT") {
      throw new Error(`${label} does not exist: ${resolved}`);
    }
    throw error;
  }
  if (st.isSymbolicLink()) {
    throw new Error(`${label} must not be a symlink/junction: ${resolved}`);
  }
  if (!st.isFile()) {
    throw new Error(`${label} must be a regular file: ${resolved}`);
  }
  return resolved;
}

/** Ensure every existing path component under root is a real directory (no symlinks). */
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

/**
 * @returns {{
 *   appData: string,
 *   jsaddons: string,
 *   addonDir: string,
 *   publishXml: string,
 *   stateFile: string,
 * }}
 */
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
  // create kingsoft/wps/jsaddons chain without following links
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
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

export function stagingDir(jsaddons) {
  return path.join(jsaddons, `${STAGING_PREFIX}${randomToken()}`);
}

export function prevAddonDir(jsaddons) {
  return path.join(jsaddons, `${PREV_PREFIX}${randomToken()}`);
}

export function ownPublishBackupPath(jsaddons) {
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  return path.join(jsaddons, `${PUBLISH_BACKUP_PREFIX}${stamp}`);
}

/** Rotate only our backup prefix; never touch publish.xml.bak.* */
export function rotateOwnPublishBackups(jsaddons) {
  assertRealDirectory(jsaddons, "jsaddons");
  const entries = fs
    .readdirSync(jsaddons, { withFileTypes: true })
    .filter((e) => e.isFile() && e.name.startsWith(PUBLISH_BACKUP_PREFIX))
    .map((e) => e.name)
    .sort();
  while (entries.length > MAX_OWN_PUBLISH_BACKUPS) {
    const name = entries.shift();
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
  // re-check ancestry
  assertAncestryReal(path.dirname(t), p);
  fs.rmSync(t, { recursive: true, force: false });
}

export function safeRenameInside(parent, from, to) {
  assertInside(parent, from, "rename from");
  assertInside(parent, to, "rename to");
  if (fs.existsSync(to)) {
    throw new Error(`rename destination already exists: ${to}`);
  }
  fs.renameSync(from, to);
}

export function tmpdir() {
  return os.tmpdir();
}
