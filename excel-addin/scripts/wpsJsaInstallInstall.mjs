/**
 * Transactional WPS JSA install (install-time Node CLI only).
 */
import { createHash } from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { createWpsPackage } from "./package-wps-jsa.mjs";
import {
  hashMapToObject,
  packageDigest,
  validateWpsPackageDir,
} from "./wpsJsaInstallValidate.mjs";
import {
  emptyPublish,
  upsertOwnPlugin,
  writePublishXmlAtomic,
  restorePublishBytes,
} from "./wpsJsaInstallPublish.mjs";
import {
  writeStateAtomic,
  restoreStateBytes,
  hashAddonTree,
} from "./wpsJsaInstallState.mjs";
import {
  assertInside,
  assertRealDirectory,
  assertRealFile,
  ensureJsaddonsDir,
  listActiveTempNames,
  lstatIfPresent,
  preflightExistingSurface,
  preflightOwnPrefixedEntries,
  prevAddonDir,
  resolveAppDataRoot,
  resolveJsaddonsLayout,
  safeRenameInside,
  safeRmInside,
  stagingDir,
} from "./wpsJsaInstallPaths.mjs";
import { WPS_ADDON_DIRECTORY, WPS_ADDON_NAME } from "./wpsJsaPackage.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const defaultRoot = path.resolve(__dirname, "..");

function copyDirReal(src, dest, destRoot) {
  assertRealDirectory(src, "copy source");
  assertInside(destRoot, dest, "copy dest");
  fs.mkdirSync(dest, { recursive: false });
  for (const ent of fs.readdirSync(src, { withFileTypes: true })) {
    const from = path.join(src, ent.name);
    const to = path.join(dest, ent.name);
    assertInside(destRoot, to, "copy dest entry");
    if (ent.isSymbolicLink()) {
      throw new Error(`symlink forbidden in package: ${from}`);
    }
    if (ent.isDirectory()) {
      copyDirReal(from, to, destRoot);
    } else if (ent.isFile()) {
      fs.copyFileSync(from, to);
      const st = fs.lstatSync(to);
      if (st.isSymbolicLink() || !st.isFile()) {
        throw new Error(`copy produced non-regular file: ${to}`);
      }
    } else {
      throw new Error(`unsupported package entry: ${from}`);
    }
  }
}

/** Exact key set + hash re-verify of staged addon tree. */
function verifyStagedAddonExact(stageAddonDir, packageValidation) {
  const expected = new Map(
    [...packageValidation.hashes.entries()].filter(([k]) =>
      k.startsWith(`${WPS_ADDON_DIRECTORY}/`),
    ),
  );
  const actual = hashAddonTree(stageAddonDir);
  const expKeys = [...expected.keys()].sort();
  const actKeys = [...actual.keys()].sort();
  if (expKeys.length !== actKeys.length || expKeys.some((k, i) => k !== actKeys[i])) {
    throw new Error(
      `staged addon file set mismatch: expected ${expKeys.length} files, got ${actKeys.length}`,
    );
  }
  for (const [rel, exp] of expected) {
    if (actual.get(rel) !== exp) {
      throw new Error(`staged hash mismatch for ${rel}`);
    }
  }
}

function snapshotBytesIfPresent(filePath, label) {
  const st = lstatIfPresent(filePath);
  if (!st) return { existed: false, bytes: null };
  if (st.isSymbolicLink()) throw new Error(`${label} is a symlink`);
  if (!st.isFile()) throw new Error(`${label} is not a regular file`);
  return { existed: true, bytes: fs.readFileSync(filePath, "utf8") };
}

function compoundError(primary, rollbackErrors) {
  if (!rollbackErrors.length) return primary;
  const detail = rollbackErrors.map((e) => (e instanceof Error ? e.message : String(e))).join("; ");
  const err = new Error(
    `${primary instanceof Error ? primary.message : String(primary)} | rollback incomplete: ${detail}`,
  );
  err.cause = primary;
  err.rollbackErrors = rollbackErrors;
  return err;
}

/**
 * @param {{
 *   packageDir?: string|null,
 *   appData?: string|null,
 *   gitSha?: string|null,
 *   rootDir?: string,
 *   skipBuild?: boolean,
 *   platform?: string,
 *   env?: NodeJS.ProcessEnv,
 *   failAfter?: string|null,
 *   afterValidate?: (packageDir: string) => void,
 * }} opts
 */
export function installWpsJsa(opts = {}) {
  const warnings = [];
  let packageDir = opts.packageDir ? path.resolve(opts.packageDir) : null;
  let built = false;

  if (!packageDir) {
    const summary = createWpsPackage({
      rootDir: opts.rootDir || defaultRoot,
      gitSha: opts.gitSha || undefined,
      skipBuild: opts.skipBuild === true,
    });
    packageDir = summary.distDir;
    built = true;
  }

  let validated = validateWpsPackageDir(packageDir);
  if (typeof opts.afterValidate === "function") {
    opts.afterValidate(packageDir);
  }
  // Re-validate after optional seam so package TOCTOU fails before appData mutation.
  validated = validateWpsPackageDir(packageDir);

  const appData = resolveAppDataRoot({
    appData: opts.appData,
    platform: opts.platform,
    env: opts.env,
  });
  const layout = resolveJsaddonsLayout(appData);
  ensureJsaddonsDir(layout.jsaddons, layout.appData);
  preflightExistingSurface(layout);
  preflightOwnPrefixedEntries(layout.jsaddons);

  const oldPublish = snapshotBytesIfPresent(layout.publishXml, "publish.xml");
  const oldState = snapshotBytesIfPresent(layout.stateFile, "state");
  const hadOldAddon = Boolean(lstatIfPresent(layout.addonDir));
  if (hadOldAddon) assertRealDirectory(layout.addonDir, "existing addon");

  let currentPublish = emptyPublish();
  if (oldPublish.existed) currentPublish = oldPublish.bytes;
  const merged = upsertOwnPlugin(currentPublish);
  warnings.push(...merged.warnings);

  const stageRoot = stagingDir(layout.jsaddons);
  const stageAddon = path.join(stageRoot, WPS_ADDON_DIRECTORY);
  let prevDir = null;
  let swapped = false;
  let publishCommitted = false;
  const failAfter = opts.failAfter || null;

  function runFail(name) {
    if (failAfter === name) throw new Error(`failpoint:${name}`);
  }

  function rollback(primary) {
    const rbErrors = [];
    try {
      if (swapped) {
        if (lstatIfPresent(layout.addonDir)) {
          safeRmInside(layout.jsaddons, layout.addonDir);
        }
        if (prevDir && lstatIfPresent(prevDir)) {
          safeRenameInside(layout.jsaddons, prevDir, layout.addonDir);
          prevDir = null;
        }
      } else if (prevDir && lstatIfPresent(prevDir) && !lstatIfPresent(layout.addonDir)) {
        safeRenameInside(layout.jsaddons, prevDir, layout.addonDir);
        prevDir = null;
      }
    } catch (e) {
      rbErrors.push(e);
    }
    try {
      restorePublishBytes(
        layout.jsaddons,
        layout.publishXml,
        oldPublish.bytes,
        oldPublish.existed,
      );
    } catch (e) {
      rbErrors.push(e);
    }
    try {
      restoreStateBytes(
        layout.jsaddons,
        layout.stateFile,
        oldState.bytes,
        oldState.existed,
      );
    } catch (e) {
      rbErrors.push(e);
    }
    try {
      if (lstatIfPresent(stageRoot)) safeRmInside(layout.jsaddons, stageRoot);
    } catch (e) {
      rbErrors.push(e);
    }
    try {
      if (prevDir && lstatIfPresent(prevDir)) safeRmInside(layout.jsaddons, prevDir);
    } catch (e) {
      rbErrors.push(e);
    }
    throw compoundError(primary, rbErrors);
  }

  try {
    copyDirReal(validated.addonDir, stageAddon, layout.jsaddons);
    // Re-validate after copy against original validated hashes (exact set)
    // and re-read staged tree so post-validate package extras fail.
    verifyStagedAddonExact(stageAddon, validated);

    if (hadOldAddon) {
      prevDir = prevAddonDir(layout.jsaddons);
      fs.rmdirSync(prevDir);
      safeRenameInside(layout.jsaddons, layout.addonDir, prevDir);
    }

    safeRenameInside(layout.jsaddons, stageAddon, layout.addonDir);
    swapped = true;
    try {
      if (lstatIfPresent(stageRoot)) safeRmInside(layout.jsaddons, stageRoot);
    } catch {
      /* best effort empty stage root */
    }
    runFail("addon-swap");

    writePublishXmlAtomic(layout.jsaddons, layout.publishXml, merged.xml, {
      failBeforeRename: () => runFail("publish-write"),
      failAfterCommit: () => runFail("publish-write-after"),
      collectRotateWarning: (msg) => warnings.push(`publish backup rotate: ${msg}`),
    });
    publishCommitted = true;

    const addonHashes = new Map(
      [...validated.hashes.entries()].filter(([k]) =>
        k.startsWith(`${WPS_ADDON_DIRECTORY}/`),
      ),
    );
    const state = {
      schemaVersion: 1,
      addonName: WPS_ADDON_NAME,
      addonDirectory: WPS_ADDON_DIRECTORY,
      installedAt: new Date().toISOString(),
      packageVersion: validated.buildInfo.packageVersion,
      gitSha: validated.buildInfo.gitSha,
      publishUrl: validated.buildInfo.publishUrl,
      packageDigest: packageDigest(validated.hashes),
      fileHashes: hashMapToObject(addonHashes),
      restartRequired: true,
      builtPackage: built,
    };

    writeStateAtomic(layout.jsaddons, layout.stateFile, state, {
      failBeforeRename: () => runFail("state-write"),
      failAfterCommit: () => runFail("state-write-after"),
    });

    // success: cleanup previous addon backup
    if (prevDir && lstatIfPresent(prevDir)) {
      try {
        safeRmInside(layout.jsaddons, prevDir);
        prevDir = null;
      } catch (error) {
        warnings.push(
          `install committed but previous addon backup cleanup failed: ${
            error instanceof Error ? error.message : String(error)
          }`,
        );
      }
    }

    return {
      ok: true,
      action: "install",
      installed: true,
      addonDirectory: WPS_ADDON_DIRECTORY,
      packageVersion: state.packageVersion,
      gitSha: state.gitSha,
      publishXml: layout.publishXml,
      addonDir: layout.addonDir,
      stateFile: layout.stateFile,
      restartRequired: true,
      message:
        "Installed. Fully quit and restart WPS before loading the add-in. This tool does not start or stop WPS.",
      warnings,
      packageDir,
      activeTemps: listActiveTempNames(layout.jsaddons),
    };
  } catch (error) {
    if (error && error.rollbackErrors) throw error;
    throw rollback(error);
  }
}
