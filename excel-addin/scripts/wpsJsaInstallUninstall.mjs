/**
 * Transactional uninstall: rename addon+state aside, update publish, then cleanup.
 * State/addon never permanently deleted before commit point.
 */
import fs from "node:fs";
import path from "node:path";
import {
  assertInside,
  assertRealDirectory,
  assertRealFile,
  listActiveTempNames,
  lstatIfPresent,
  preflightExistingSurface,
  preflightOwnPrefixedEntries,
  PREV_PREFIX,
  resolveAppDataRoot,
  resolveJsaddonsLayout,
  reserveExclusiveDirPath,
  reserveExclusivePath,
  safeRenameInside,
  safeRmInside,
  STATE_FILE_NAME,
  TMP_PREFIX,
} from "./wpsJsaInstallPaths.mjs";
import {
  parseJspluginsDocument,
  removeOwnPlugin,
  restorePublishBytes,
  writePublishXmlAtomic,
} from "./wpsJsaInstallPublish.mjs";
import { WPS_ADDON_DIRECTORY, WPS_ADDON_NAME } from "./wpsJsaPackage.mjs";

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
 *   appData?: string|null,
 *   platform?: string,
 *   env?: NodeJS.ProcessEnv,
 *   failAfter?: 'publish-before'|'publish-after'|'addon-move'|'state-move'|'state'|null,
 * }} opts
 */
export function uninstallWpsJsa(opts = {}) {
  const warnings = [];
  const appData = resolveAppDataRoot({
    appData: opts.appData,
    platform: opts.platform,
    env: opts.env,
  });
  const layout = resolveJsaddonsLayout(appData);

  if (!lstatIfPresent(layout.jsaddons)) {
    return {
      ok: true,
      action: "uninstall",
      removed: false,
      message: "jsaddons not present; nothing to uninstall",
      warnings,
      restartRequired: true,
    };
  }
  assertRealDirectory(layout.jsaddons, "jsaddons");
  preflightExistingSurface(layout);
  preflightOwnPrefixedEntries(layout.jsaddons);

  const pubSt = lstatIfPresent(layout.publishXml);
  let ownInPublish = false;
  let publishBytes = null;
  if (pubSt) {
    assertRealFile(layout.publishXml, "publish.xml");
    publishBytes = fs.readFileSync(layout.publishXml, "utf8");
    const parsed = parseJspluginsDocument(publishBytes);
    warnings.push(...parsed.warnings);
    ownInPublish = parsed.plugins.some((p) => p.attrs.name === WPS_ADDON_NAME);
  }

  const addonSt = lstatIfPresent(layout.addonDir);
  let hasAddon = false;
  if (addonSt) {
    assertRealDirectory(layout.addonDir, "addonDir");
    if (path.basename(layout.addonDir) !== WPS_ADDON_DIRECTORY) {
      throw new Error("refusing to remove unexpected addon directory name");
    }
    hasAddon = true;
  }

  const stateSt = lstatIfPresent(layout.stateFile);
  let hasState = false;
  if (stateSt) {
    assertRealFile(layout.stateFile, "state");
    if (path.basename(layout.stateFile) !== STATE_FILE_NAME) {
      throw new Error("state basename mismatch");
    }
    hasState = true;
  }

  if (!ownInPublish && !hasAddon && !hasState) {
    return {
      ok: true,
      action: "uninstall",
      removed: false,
      message: "Nothing of this add-in is installed; foreign plugins left untouched",
      warnings,
      restartRequired: true,
      publishXml: layout.publishXml,
      addonDir: layout.addonDir,
    };
  }

  const failAfter = opts.failAfter || null;
  let publishChanged = false;
  let addonBackup = null;
  let stateBackup = null;
  let committed = false;

  function rollback(primary) {
    const rbErrors = [];
    try {
      if (publishChanged) {
        restorePublishBytes(
          layout.jsaddons,
          layout.publishXml,
          publishBytes,
          Boolean(pubSt),
        );
      }
    } catch (e) {
      rbErrors.push(e);
    }
    try {
      if (addonBackup && lstatIfPresent(addonBackup) && !lstatIfPresent(layout.addonDir)) {
        safeRenameInside(layout.jsaddons, addonBackup, layout.addonDir);
        addonBackup = null;
      }
    } catch (e) {
      rbErrors.push(e);
    }
    try {
      if (stateBackup && lstatIfPresent(stateBackup) && !lstatIfPresent(layout.stateFile)) {
        safeRenameInside(layout.jsaddons, stateBackup, layout.stateFile);
        stateBackup = null;
      }
    } catch (e) {
      rbErrors.push(e);
    }
    throw compoundError(primary, rbErrors);
  }

  try {
    // 1) Move addon aside (recoverable until commit)
    if (hasAddon) {
      const dest = reserveExclusiveDirPath(layout.jsaddons, PREV_PREFIX);
      safeRenameInside(layout.jsaddons, layout.addonDir, dest);
      addonBackup = dest;
      if (failAfter === "addon-move") throw new Error("failpoint:addon-move");
    }

    // 2) Move state aside (recoverable) — never permanent delete before commit
    if (hasState) {
      const dest = reserveExclusivePath(layout.jsaddons, `${TMP_PREFIX}state-bak-`);
      safeRenameInside(layout.jsaddons, layout.stateFile, dest);
      stateBackup = dest;
      if (failAfter === "state-move" || failAfter === "state") {
        throw new Error(failAfter === "state" ? "failpoint:state" : "failpoint:state-move");
      }
    }

    // 3) Publish update
    if (ownInPublish) {
      const removed = removeOwnPlugin(publishBytes);
      warnings.push(...removed.warnings);
      if (!removed.removed) throw new Error("internal: expected own plugin removal");
      if (failAfter === "publish-before") throw new Error("failpoint:publish-before");
      // Mark before write so post-commit failpoints still restore.
      publishChanged = true;
      writePublishXmlAtomic(layout.jsaddons, layout.publishXml, removed.xml, {
        failBeforeRename: () => {
          if (failAfter === "publish-before") throw new Error("failpoint:publish-before");
        },
        failAfterCommit: () => {
          if (failAfter === "publish-after") throw new Error("failpoint:publish-after");
        },
        collectRotateWarning: (msg) => warnings.push(`publish backup rotate: ${msg}`),
      });
    }

    // COMMIT POINT
    committed = true;

    // 4) Best-effort cleanup of recoverable backups
    if (addonBackup && lstatIfPresent(addonBackup)) {
      try {
        safeRmInside(layout.jsaddons, addonBackup);
        addonBackup = null;
      } catch (error) {
        warnings.push(
          `uninstall committed but addon backup cleanup failed: ${
            error instanceof Error ? error.message : String(error)
          }`,
        );
      }
    }
    if (stateBackup && lstatIfPresent(stateBackup)) {
      try {
        const st = lstatIfPresent(stateBackup);
        if (!st || st.isSymbolicLink() || !st.isFile()) {
          throw new Error(`state backup not a regular file: ${stateBackup}`);
        }
        assertInside(layout.jsaddons, stateBackup, "state backup");
        fs.unlinkSync(stateBackup);
        stateBackup = null;
      } catch (error) {
        warnings.push(
          `uninstall committed but state backup cleanup failed: ${
            error instanceof Error ? error.message : String(error)
          }`,
        );
      }
    }

    return {
      ok: true,
      action: "uninstall",
      removed: true,
      message:
        "Own add-in entry/directory/state removed. Fully restart WPS if it was running. Foreign plugins and legacy dirs were left untouched.",
      warnings,
      restartRequired: true,
      publishXml: layout.publishXml,
      addonDir: layout.addonDir,
      activeTemps: listActiveTempNames(layout.jsaddons),
    };
  } catch (error) {
    if (committed) throw error;
    if (error && error.rollbackErrors) throw error;
    throw rollback(error);
  }
}
