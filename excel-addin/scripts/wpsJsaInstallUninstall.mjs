/**
 * Transactional uninstall of own WPS JSA entry only.
 */
import fs from "node:fs";
import path from "node:path";
import {
  assertRealDirectory,
  assertRealFile,
  preflightExistingSurface,
  resolveAppDataRoot,
  resolveJsaddonsLayout,
  safeRmInside,
  exclusiveTempDir,
  PREV_PREFIX,
  STATE_FILE_NAME,
} from "./wpsJsaInstallPaths.mjs";
import {
  parseJspluginsDocument,
  removeOwnPlugin,
  restorePublishBytes,
  writePublishXmlAtomic,
} from "./wpsJsaInstallPublish.mjs";
import { restoreStateBytes } from "./wpsJsaInstallState.mjs";
import { WPS_ADDON_DIRECTORY, WPS_ADDON_NAME } from "./wpsJsaPackage.mjs";

/**
 * @param {{
 *   appData?: string|null,
 *   platform?: string,
 *   env?: NodeJS.ProcessEnv,
 *   failAfter?: 'publish'|'addon'|'state'|null,
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

  if (!fs.existsSync(layout.jsaddons)) {
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

  const hasPublish = fs.existsSync(layout.publishXml);
  let ownInPublish = false;
  let publishBytes = null;
  if (hasPublish) {
    assertRealFile(layout.publishXml, "publish.xml");
    publishBytes = fs.readFileSync(layout.publishXml, "utf8");
    const parsed = parseJspluginsDocument(publishBytes);
    warnings.push(...parsed.warnings);
    ownInPublish = parsed.plugins.some((p) => p.attrs.name === WPS_ADDON_NAME);
  }

  const hasAddon =
    fs.existsSync(layout.addonDir) &&
    (() => {
      assertRealDirectory(layout.addonDir, "addonDir");
      if (path.basename(layout.addonDir) !== WPS_ADDON_DIRECTORY) {
        throw new Error("refusing to remove unexpected addon directory name");
      }
      return true;
    })();

  const hasState = fs.existsSync(layout.stateFile);
  if (hasState) {
    assertRealFile(layout.stateFile, "state");
    if (path.basename(layout.stateFile) !== STATE_FILE_NAME) {
      throw new Error("state basename mismatch");
    }
  }

  if (!ownInPublish && !hasAddon && !hasState) {
    // Do not rewrite publish even if whitespace differs
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
  let addonMovedTo = null;
  let stateRemoved = false;
  const oldStateBytes = hasState ? fs.readFileSync(layout.stateFile, "utf8") : null;

  function rollback() {
    try {
      if (publishChanged) {
        restorePublishBytes(layout.jsaddons, layout.publishXml, publishBytes, hasPublish);
      }
    } catch {
      /* best effort */
    }
    try {
      if (addonMovedTo && fs.existsSync(addonMovedTo) && !fs.existsSync(layout.addonDir)) {
        fs.renameSync(addonMovedTo, layout.addonDir);
      }
    } catch {
      /* best effort */
    }
    try {
      if (stateRemoved && oldStateBytes != null) {
        restoreStateBytes(layout.jsaddons, layout.stateFile, oldStateBytes, true);
      }
    } catch {
      /* best effort */
    }
  }

  try {
    if (ownInPublish) {
      const removed = removeOwnPlugin(publishBytes);
      warnings.push(...removed.warnings);
      if (!removed.removed) {
        throw new Error("internal: expected own plugin removal");
      }
      writePublishXmlAtomic(layout.jsaddons, layout.publishXml, removed.xml, {
        failBeforeRename: () => {
          if (failAfter === "publish") throw new Error("failpoint:publish");
        },
      });
      publishChanged = true;
    }

    if (hasAddon) {
      // move aside then delete for recoverability mid-flight
      const trash = exclusiveTempDir(layout.jsaddons, PREV_PREFIX);
      fs.rmdirSync(trash);
      fs.renameSync(layout.addonDir, trash);
      addonMovedTo = trash;
      if (failAfter === "addon") throw new Error("failpoint:addon");
      safeRmInside(layout.jsaddons, trash);
      addonMovedTo = null;
    }

    if (hasState) {
      if (failAfter === "state") throw new Error("failpoint:state");
      fs.unlinkSync(layout.stateFile);
      stateRemoved = true;
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
    };
  } catch (error) {
    rollback();
    throw error;
  }
}
