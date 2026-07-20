/**
 * Install state file schema + atomic IO (install-time only).
 */
import { createHash } from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import {
  assertInside,
  assertRealFile,
  exclusiveTempFile,
  STATE_FILE_NAME,
  TMP_PREFIX,
} from "./wpsJsaInstallPaths.mjs";
import { assertSafeAddonHashKey } from "./wpsJsaInstallValidate.mjs";
import {
  WPS_ADDON_DIRECTORY,
  WPS_ADDON_NAME,
  WPS_PUBLISH_URL,
} from "./wpsJsaPackage.mjs";

const SAFE_VERSION_RE = /^\d+\.\d+\.\d+(?:\.\d+)?$/;
const SAFE_SHA_RE = /^(?:unknown|[0-9a-f]{7,64})$/i;

/**
 * @returns {{ ok: true, state: object } | { ok: false, reason: string, state?: any }}
 */
export function validateInstallState(state) {
  if (state == null || typeof state !== "object" || Array.isArray(state)) {
    return { ok: false, reason: "state-invalid-type" };
  }
  if (state.schemaVersion !== 1) {
    return { ok: false, reason: "state-schemaVersion" };
  }
  if (state.addonName !== WPS_ADDON_NAME) {
    return { ok: false, reason: "state-addonName" };
  }
  if (state.addonDirectory !== WPS_ADDON_DIRECTORY) {
    return { ok: false, reason: "state-addonDirectory" };
  }
  if (state.publishUrl !== WPS_PUBLISH_URL) {
    return { ok: false, reason: "state-publishUrl" };
  }
  if (!SAFE_VERSION_RE.test(String(state.packageVersion ?? ""))) {
    return { ok: false, reason: "state-packageVersion" };
  }
  if (!SAFE_SHA_RE.test(String(state.gitSha ?? ""))) {
    return { ok: false, reason: "state-gitSha" };
  }
  if (state.restartRequired !== true) {
    return { ok: false, reason: "state-restartRequired" };
  }
  if (typeof state.packageDigest !== "string" || !/^[0-9a-f]{64}$/i.test(state.packageDigest)) {
    return { ok: false, reason: "state-packageDigest" };
  }
  if (
    state.fileHashes == null ||
    typeof state.fileHashes !== "object" ||
    Array.isArray(state.fileHashes)
  ) {
    return { ok: false, reason: "state-fileHashes-type" };
  }
  const keys = Object.keys(state.fileHashes);
  if (keys.length === 0) {
    return { ok: false, reason: "state-fileHashes-empty" };
  }
  for (const key of keys) {
    try {
      assertSafeAddonHashKey(key);
    } catch {
      return { ok: false, reason: `state-fileHashes-key:${key}` };
    }
    const h = state.fileHashes[key];
    if (typeof h !== "string" || !/^[0-9a-f]{64}$/i.test(h)) {
      return { ok: false, reason: `state-fileHashes-value:${key}` };
    }
  }
  if (state.installedAt != null && typeof state.installedAt !== "string") {
    return { ok: false, reason: "state-installedAt" };
  }
  return { ok: true, state };
}

export function readStateFile(statePath) {
  if (!fs.existsSync(statePath)) return { present: false };
  try {
    assertRealFile(statePath, "state");
  } catch (error) {
    return {
      present: true,
      invalid: true,
      reason: error instanceof Error ? error.message : String(error),
    };
  }
  if (path.basename(statePath) !== STATE_FILE_NAME) {
    return { present: true, invalid: true, reason: "state-basename" };
  }
  let parsed;
  try {
    parsed = JSON.parse(fs.readFileSync(statePath, "utf8"));
  } catch {
    return { present: true, invalid: true, reason: "state-json" };
  }
  const v = validateInstallState(parsed);
  if (!v.ok) {
    return { present: true, invalid: true, reason: v.reason, state: parsed };
  }
  return { present: true, invalid: false, state: v.state };
}

/**
 * @returns {{ previousBytes: string|null, previousExisted: boolean }}
 */
export function writeStateAtomic(jsaddons, statePath, state, opts = {}) {
  assertInside(jsaddons, statePath, "state");
  if (path.basename(statePath) !== STATE_FILE_NAME) {
    throw new Error("state basename mismatch");
  }
  let previousBytes = null;
  let previousExisted = false;
  if (fs.existsSync(statePath)) {
    assertRealFile(statePath, "state");
    previousExisted = true;
    previousBytes = fs.readFileSync(statePath, "utf8");
  }
  const tmp = exclusiveTempFile(jsaddons, `${TMP_PREFIX}state-`);
  const body = `${JSON.stringify(state, null, 2)}\n`;
  try {
    if (typeof opts.failBeforeWrite === "function") opts.failBeforeWrite();
    fs.writeFileSync(tmp, body, "utf8");
    const fd = fs.openSync(tmp, "r+");
    try {
      fs.fsyncSync(fd);
    } finally {
      fs.closeSync(fd);
    }
    if (typeof opts.failBeforeRename === "function") opts.failBeforeRename();
    fs.renameSync(tmp, statePath);
    if (typeof opts.failAfterCommit === "function") opts.failAfterCommit();
  } catch (error) {
    try {
      if (fs.existsSync(tmp)) fs.unlinkSync(tmp);
    } catch {
      /* ignore */
    }
    throw error;
  }
  return { previousBytes, previousExisted };
}

export function restoreStateBytes(jsaddons, statePath, previousBytes, previousExisted) {
  assertInside(jsaddons, statePath, "state");
  if (!previousExisted) {
    if (fs.existsSync(statePath)) {
      const st = fs.lstatSync(statePath);
      if (st.isSymbolicLink()) throw new Error("cannot restore over symlink state");
      fs.unlinkSync(statePath);
    }
    return;
  }
  fs.writeFileSync(statePath, previousBytes, "utf8");
}

/** Enumerate real files under addonDir → Map of WPS_ADDON_DIRECTORY/rel -> sha256 */
export function hashAddonTree(addonDir) {
  assertInside(path.dirname(addonDir), addonDir, "addonDir");
  const out = new Map();
  function walk(dir, relParts) {
    const ents = fs.readdirSync(dir, { withFileTypes: true });
    for (const ent of ents) {
      const abs = path.join(dir, ent.name);
      const st = fs.lstatSync(abs);
      if (st.isSymbolicLink()) {
        throw new Error(`symlink forbidden in installed addon: ${abs}`);
      }
      if (ent.isDirectory()) {
        if (!st.isDirectory()) throw new Error(`non-directory entry: ${abs}`);
        walk(abs, [...relParts, ent.name]);
      } else if (ent.isFile()) {
        if (!st.isFile()) throw new Error(`non-regular file: ${abs}`);
        const rel = [WPS_ADDON_DIRECTORY, ...relParts, ent.name].join("/");
        assertSafeAddonHashKey(rel);
        const hash = createHash("sha256").update(fs.readFileSync(abs)).digest("hex");
        out.set(rel, hash);
      } else {
        throw new Error(`unsupported addon entry: ${abs}`);
      }
    }
  }
  walk(addonDir, []);
  return out;
}
