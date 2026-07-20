/**
 * Read-only inspection and post-commit cleanup of Phase56–58 kebab-case layout.
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
  WPS_ADDON_NAME,
} from "./wpsJsaPackage.mjs";

/**
 * @param {{ jsaddons: string, legacyAddonDir: string, stateFile: string }} layout
 * @returns {{
 *   present: boolean,
 *   verified: boolean,
 *   wouldRemove: boolean,
 *   path: string,
 *   reason?: string,
 *   warning?: string,
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
  } catch (error) {
    result.reason = "legacy-path-escape";
    result.warning =
      "legacy own directory path is unsafe; left untouched";
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
    result.warning =
      "legacy own path is not a directory; left untouched";
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
  } catch (error) {
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
  } catch (error) {
    result.reason = "legacy-tree-error";
    result.warning =
      "legacy own directory tree cannot be hashed safely; not removed";
    return result;
  }

  const expected = validated.state.fileHashes;
  const expectedKeys = Object.keys(expected).sort();
  const actualKeys = [...actual.keys()].sort();
  if (expectedKeys.length !== actualKeys.length) {
    result.reason = "legacy-hash-count";
    result.warning =
      "legacy own directory hash surface drifts from state; not removed";
    return result;
  }
  for (let i = 0; i < expectedKeys.length; i += 1) {
    const k = expectedKeys[i];
    if (k !== actualKeys[i] || actual.get(k) !== expected[k]) {
      result.reason = "legacy-hash-mismatch";
      result.warning =
        "legacy own directory content does not match install state; not removed";
      return result;
    }
  }

  if (validated.state.addonName !== WPS_ADDON_NAME) {
    result.reason = "legacy-name-mismatch";
    result.warning = "legacy state addonName mismatch; not removed";
    return result;
  }

  result.verified = true;
  result.wouldRemove = true;
  return result;
}

/**
 * Best-effort post-commit removal of a previously verified legacy directory.
 * Never throws for cleanup failure — returns warning string or null.
 * @param {{ jsaddons: string, legacyAddonDir: string }} layout
 * @param {{ verified: boolean, present: boolean, path: string }} legacy
 */
export function removeVerifiedLegacyOwnAddon(layout, legacy) {
  if (!legacy || !legacy.verified || !legacy.present) return null;
  const target = legacy.path || layout.legacyAddonDir;
  try {
    const st = lstatIfPresent(target);
    if (!st) return null;
    if (st.isSymbolicLink() || !st.isDirectory()) {
      return "legacy own directory changed after commit; left untouched";
    }
    assertInside(layout.jsaddons, target, "legacy addon cleanup");
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
