/**
 * Transactional WPS JSA install (install-time Node CLI only).
 * Consumes shared planWpsJsaInstall; --dry-run never mutates AppData.
 */
import fs from "node:fs";
import path from "node:path";
import {
  packageDigest,
  validateWpsPackageDir,
} from "./wpsJsaInstallValidate.mjs";
import {
  writePublishXmlAtomic,
  restorePublishBytes,
  upsertOwnPlugin,
  emptyPublish,
} from "./wpsJsaInstallPublish.mjs";
import {
  writeStateAtomic,
  restoreStateBytes,
  hashAddonTree,
} from "./wpsJsaInstallState.mjs";
import {
  formatDryRunResult,
  planWpsJsaInstall,
  resolveAndValidatePackage,
} from "./wpsJsaInstallPlan.mjs";
import {
  projectPublicPluginNames,
  projectPublicWarnings,
} from "./wpsJsaInstallPublicNames.mjs";
import { removeVerifiedLegacyOwnAddon } from "./wpsJsaInstallLegacy.mjs";
import {
  assertInside,
  assertRealDirectory,
  ensureJsaddonsDir,
  listActiveTempNames,
  lstatIfPresent,
  preflightExistingSurface,
  preflightOwnPrefixedEntries,
  prevAddonDir,
  safeRenameInside,
  safeRmInside,
  stagingDir,
} from "./wpsJsaInstallPaths.mjs";
import {
  LEGACY_OWN_ADDON_DIRECTORY,
  WPS_ADDON_DIRECTORY,
  WPS_ADDON_NAME,
} from "./wpsJsaPackage.mjs";
import { hashMapToObject } from "./wpsJsaInstallValidate.mjs";

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
  const detail = rollbackErrors
    .map((e) => (e instanceof Error ? e.message : String(e)))
    .join("; ");
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
 *   dryRun?: boolean,
 *   failAfter?: string|null,
 *   afterValidate?: (packageDir: string) => void,
 *   createWpsPackage?: Function,
 * }} opts
 */
export function installWpsJsa(opts = {}) {
  if (opts.dryRun === true) {
    const plan = planWpsJsaInstall(opts);
    return formatDryRunResult(plan);
  }

  // Shared package resolution + plan (read-only AppData inspection)
  const packageResolution = resolveAndValidatePackage(opts);
  const plan = planWpsJsaInstall({
    ...opts,
    packageResolution,
  });

  const warnings = [...plan.warnings];
  const layout = plan.layout;
  const appData = plan.appData;

  // Mutations begin: ensure dirs then TOCTOU revalidate package + re-preflight
  ensureJsaddonsDir(layout.jsaddons, layout.appData);
  const validated = validateWpsPackageDir(packageResolution.packageDir);
  preflightExistingSurface(layout);
  preflightOwnPrefixedEntries(layout.jsaddons);

  const oldPublish = snapshotBytesIfPresent(layout.publishXml, "publish.xml");
  const oldState = snapshotBytesIfPresent(layout.stateFile, "state");
  const hadOldAddon = Boolean(lstatIfPresent(layout.addonDir));
  if (hadOldAddon) assertRealDirectory(layout.addonDir, "existing addon");

  // Re-plan publish merge from live bytes after TOCTOU
  let currentPublish = emptyPublish();
  if (oldPublish.existed) currentPublish = oldPublish.bytes;
  const merged = upsertOwnPlugin(currentPublish);
  warnings.push(...merged.warnings);

  const stageRoot = stagingDir(layout.jsaddons);
  const stageAddon = path.join(stageRoot, WPS_ADDON_DIRECTORY);
  let prevDir = null;
  let swapped = false;
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
      /* best effort */
    }
    runFail("addon-swap");

    writePublishXmlAtomic(layout.jsaddons, layout.publishXml, merged.xml, {
      failBeforeRename: () => runFail("publish-write"),
      failAfterCommit: () => runFail("publish-write-after"),
      collectRotateWarning: (msg) => warnings.push(`publish backup rotate: ${msg}`),
    });

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
      builtPackage: packageResolution.built,
    };
    if (plan.legacyOwnAddonVerified) {
      state.migratedFromAddonDirectory = LEGACY_OWN_ADDON_DIRECTORY;
    }

    writeStateAtomic(layout.jsaddons, layout.stateFile, state, {
      failBeforeRename: () => runFail("state-write"),
      failAfterCommit: () => runFail("state-write-after"),
    });

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

    // Post-commit only: remove verified Phase56–58 kebab-case directory.
    const legacyCleanupWarning = removeVerifiedLegacyOwnAddon(layout, plan.legacyOwn);
    if (legacyCleanupWarning) warnings.push(legacyCleanupWarning);

    return {
      ok: true,
      action: "install",
      dryRun: false,
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
      warnings: projectPublicWarnings(warnings),
      packageDir: packageResolution.packageDir,
      activeTemps: listActiveTempNames(layout.jsaddons),
      // plan consistency fields (actual)
      wouldCreateJsaddons: plan.wouldCreateJsaddons,
      wouldReplaceAddon: plan.wouldReplaceAddon,
      wouldCreatePublish: plan.wouldCreatePublish,
      wouldUpdatePublish: plan.wouldUpdatePublish,
      wouldWriteState: true,
      existingOwnEntry: plan.existingOwnEntry,
      legacyOwnAddonPresent: plan.legacyOwnAddonPresent === true,
      legacyOwnAddonVerified: plan.legacyOwnAddonVerified === true,
      wouldRemoveLegacyOwnAddon: plan.wouldRemoveLegacyOwnAddon === true,
      migratedFromAddonDirectory: state.migratedFromAddonDirectory || null,
      preservedPluginNames: projectPublicPluginNames(plan.preservedPluginNames || []),
    };
  } catch (error) {
    if (error && error.rollbackErrors) throw error;
    throw rollback(error);
  }
}
