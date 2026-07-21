/**
 * Read-only inspection and post-commit cleanup of Phase56–58 kebab-case layout.
 * Cleanup must re-verify against an immutable expectedSnapshot from plan time —
 * never re-read the live state file (install may have already replaced it).
 */
import fs from "node:fs";
import path from "node:path";
import {
  assertInside,
  assertRealDirectory,
  lstatIfPresent,
  safeRmInside,
} from "./wpsJsaInstallPaths.mjs";
import {
  hashAddonTree,
  validateLegacyOwnInstallState,
} from "./wpsJsaInstallState.mjs";
import {
  LEGACY_OWN_ADDON_DIRECTORY,
  LEGACY_OWN_PUBLISH_URL,
  WPS_ADDON_NAME,
} from "./wpsJsaPackage.mjs";

/**
 * @param {unknown} snap
 * @returns {snap is {
 *   addonName: string,
 *   addonDirectory: string,
 *   publishUrl: string,
 *   packageVersion: string,
 *   gitSha: string,
 *   packageDigest: string,
 *   fileHashes: Record<string, string>,
 * }}
 */
export function isLegacyExpectedSnapshot(snap) {
  if (snap == null || typeof snap !== "object" || Array.isArray(snap)) return false;
  const s = /** @type {Record<string, unknown>} */ (snap);
  if (s.addonName !== WPS_ADDON_NAME) return false;
  if (s.addonDirectory !== LEGACY_OWN_ADDON_DIRECTORY) return false;
  if (s.publishUrl !== LEGACY_OWN_PUBLISH_URL) return false;
  if (typeof s.packageVersion !== "string" || s.packageVersion === "") return false;
  if (typeof s.gitSha !== "string" || s.gitSha === "") return false;
  if (typeof s.packageDigest !== "string" || !/^[0-9a-f]{64}$/i.test(s.packageDigest)) {
    return false;
  }
  if (s.fileHashes == null || typeof s.fileHashes !== "object" || Array.isArray(s.fileHashes)) {
    return false;
  }
  const keys = Object.keys(/** @type {object} */ (s.fileHashes));
  if (keys.length === 0) return false;
  for (const key of keys) {
    if (typeof key !== "string" || !key.startsWith(`${LEGACY_OWN_ADDON_DIRECTORY}/`)) {
      return false;
    }
    const h = /** @type {Record<string, unknown>} */ (s.fileHashes)[key];
    if (typeof h !== "string" || !/^[0-9a-f]{64}$/i.test(h)) return false;
  }
  return true;
}

/**
 * Freeze a plain, immutable-enough snapshot of verified legacy state contract + hashes.
 * @param {object} state validated legacy install state
 */
function freezeExpectedSnapshot(state) {
  /** @type {Record<string, string>} */
  const fileHashes = {};
  for (const key of Object.keys(state.fileHashes).sort()) {
    fileHashes[key] = String(state.fileHashes[key]);
  }
  return Object.freeze({
    addonName: state.addonName,
    addonDirectory: state.addonDirectory,
    publishUrl: state.publishUrl,
    packageVersion: state.packageVersion,
    gitSha: state.gitSha,
    packageDigest: state.packageDigest,
    fileHashes: Object.freeze(fileHashes),
  });
}

/**
 * Compare live tree hashes to a frozen expected snapshot (exact key set + SHA-256).
 * @param {Map<string, string>} actual
 * @param {Record<string, string>} expected
 */
export function legacyHashesMatchExact(actual, expected) {
  const expectedKeys = Object.keys(expected).sort();
  const actualKeys = [...actual.keys()].sort();
  if (expectedKeys.length !== actualKeys.length) return false;
  for (let i = 0; i < expectedKeys.length; i += 1) {
    const k = expectedKeys[i];
    if (k !== actualKeys[i] || actual.get(k) !== expected[k]) return false;
  }
  return true;
}

/**
 * @param {{ jsaddons: string, legacyAddonDir: string, stateFile: string }} layout
 * @returns {{
 *   present: boolean,
 *   verified: boolean,
 *   wouldRemove: boolean,
 *   path: string,
 *   reason?: string,
 *   warning?: string,
 *   expectedSnapshot?: object,
 * }}
 */
export function inspectLegacyOwnAddon(layout) {
  const legacyPath = layout.legacyAddonDir || path.join(layout.jsaddons, LEGACY_OWN_ADDON_DIRECTORY);
  const result = {
    present: false,
    verified: false,
    wouldRemove: false,
    path: legacyPath,
  };

  const st = lstatIfPresent(legacyPath);
  if (!st) return result;

  result.present = true;
  try {
    assertInside(layout.jsaddons, legacyPath, "legacy addon");
  } catch {
    result.reason = "legacy-path-escape";
    result.warning = "legacy own directory path is unsafe; left untouched";
    return result;
  }

  if (st.isSymbolicLink()) {
    result.reason = "legacy-symlink";
    result.warning =
      "legacy own directory is a symlink/junction; left untouched (not removed)";
    return result;
  }
  if (!st.isDirectory()) {
    result.reason = "legacy-not-directory";
    result.warning = "legacy own path is not a directory; left untouched";
    return result;
  }

  // Require current state file to be the legacy schema with matching hashes.
  if (!lstatIfPresent(layout.stateFile)) {
    result.reason = "legacy-state-missing";
    result.warning =
      "legacy own directory present but install state missing; not removed";
    return result;
  }

  let parsed;
  try {
    const stState = lstatIfPresent(layout.stateFile);
    if (!stState || stState.isSymbolicLink() || !stState.isFile()) {
      result.reason = "legacy-state-not-file";
      result.warning =
        "legacy own directory present but state is not a regular file; not removed";
      return result;
    }
    parsed = JSON.parse(fs.readFileSync(layout.stateFile, "utf8"));
  } catch {
    result.reason = "legacy-state-parse";
    result.warning =
      "legacy own directory present but state unreadable; not removed";
    return result;
  }

  const validated = validateLegacyOwnInstallState(parsed);
  if (!validated.ok) {
    result.reason = validated.reason || "legacy-state-invalid";
    result.warning =
      "legacy own directory present but state is not a verified legacy install; not removed";
    return result;
  }

  let actual;
  try {
    assertRealDirectory(legacyPath, "legacy addon");
    actual = hashAddonTree(legacyPath, LEGACY_OWN_ADDON_DIRECTORY);
  } catch {
    result.reason = "legacy-tree-error";
    result.warning =
      "legacy own directory tree cannot be hashed safely; not removed";
    return result;
  }

  if (!legacyHashesMatchExact(actual, validated.state.fileHashes)) {
    result.reason = "legacy-hash-mismatch";
    result.warning =
      "legacy own directory content does not match install state; not removed";
    return result;
  }

  if (validated.state.addonName !== WPS_ADDON_NAME) {
    result.reason = "legacy-name-mismatch";
    result.warning = "legacy state addonName mismatch; not removed";
    return result;
  }

  const expectedSnapshot = freezeExpectedSnapshot(validated.state);
  if (!isLegacyExpectedSnapshot(expectedSnapshot)) {
    result.reason = "legacy-snapshot-invalid";
    result.warning =
      "legacy own directory present but expected snapshot is invalid; not removed";
    return result;
  }

  result.verified = true;
  result.wouldRemove = true;
  result.expectedSnapshot = expectedSnapshot;
  return result;
}

/**
 * Re-verify live tree against plan-time expectedSnapshot, then best-effort remove.
 * Does not read stateFile (may already be the new install contract).
 * Never throws for cleanup failure — returns warning string or null.
 *
 * @param {{ jsaddons: string, legacyAddonDir?: string }} layout
 * @param {{
 *   verified?: boolean,
 *   present?: boolean,
 *   path?: string,
 *   expectedSnapshot?: object,
 * }} legacy
 * @param {{ beforeReverify?: (ctx: { target: string, layout: object, legacy: object }) => void }} [hooks]
 */
export function removeVerifiedLegacyOwnAddon(layout, legacy, hooks = {}) {
  if (!legacy || !legacy.verified || !legacy.present) return null;
  if (!isLegacyExpectedSnapshot(legacy.expectedSnapshot)) {
    return "legacy cleanup refused: missing plan-time expected snapshot; left untouched";
  }

  const target = legacy.path || layout.legacyAddonDir;
  try {
    if (typeof hooks.beforeReverify === "function") {
      hooks.beforeReverify({ target, layout, legacy });
    }

    const st = lstatIfPresent(target);
    if (!st) {
      return "legacy own directory disappeared after commit; nothing to remove";
    }
    if (st.isSymbolicLink() || !st.isDirectory()) {
      return "legacy own directory type drifted after commit; left untouched";
    }
    assertInside(layout.jsaddons, target, "legacy addon cleanup");
    if (path.basename(target) !== LEGACY_OWN_ADDON_DIRECTORY) {
      return "legacy cleanup refused: unexpected directory name";
    }

    assertRealDirectory(target, "legacy addon cleanup");
    const actual = hashAddonTree(target, LEGACY_OWN_ADDON_DIRECTORY);
    if (!legacyHashesMatchExact(actual, legacy.expectedSnapshot.fileHashes)) {
      return "legacy own directory content drifted after commit; left untouched";
    }

    // Final type recheck immediately before destructive remove.
    const st2 = lstatIfPresent(target);
    if (!st2 || st2.isSymbolicLink() || !st2.isDirectory()) {
      return "legacy own directory changed before remove; left untouched";
    }
    if (path.basename(target) !== LEGACY_OWN_ADDON_DIRECTORY) {
      return "legacy cleanup refused: unexpected directory name";
    }

    safeRmInside(layout.jsaddons, target);
    return null;
  } catch (error) {
    return `legacy own directory cleanup failed after successful install: ${
      error instanceof Error ? error.message : String(error)
    }`;
  }
}
