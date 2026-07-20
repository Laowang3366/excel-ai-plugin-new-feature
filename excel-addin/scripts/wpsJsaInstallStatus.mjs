/**
 * Honest WPS JSA install status (no symlink follow; full hash surface).
 */
import fs from "node:fs";
import {
  assertRealDirectory,
  assertRealFile,
  resolveAppDataRoot,
  resolveJsaddonsLayout,
} from "./wpsJsaInstallPaths.mjs";
import {
  ownPluginMatchesContract,
  parseJspluginsDocument,
} from "./wpsJsaInstallPublish.mjs";
import { projectPublicWarnings } from "./wpsJsaInstallPublicNames.mjs";
import { inspectLegacyOwnAddon } from "./wpsJsaInstallLegacy.mjs";
import { LEGACY_OWN_ADDON_DIRECTORY } from "./wpsJsaPackage.mjs";
import { hashAddonTree, readStateFile } from "./wpsJsaInstallState.mjs";
import { WPS_ADDON_DIRECTORY, WPS_ADDON_NAME, WPS_ENTRY_SCRIPT } from "./wpsJsaPackage.mjs";

/**
 * @param {{ appData?: string|null, platform?: string, env?: NodeJS.ProcessEnv }} opts
 */
export function statusWpsJsa(opts = {}) {
  const warnings = [];
  const drift = [];
  const appData = resolveAppDataRoot({
    appData: opts.appData,
    platform: opts.platform,
    env: opts.env,
  });
  const layout = resolveJsaddonsLayout(appData);

  const result = {
    ok: true,
    action: "status",
    installed: false,
    current: false,
    drift,
    warnings,
    restartRequired: null,
    publishXml: layout.publishXml,
    addonDir: layout.addonDir,
    stateFile: layout.stateFile,
    packageVersion: null,
    gitSha: null,
    message: "Not installed",
  };

  if (!fs.existsSync(layout.jsaddons)) {
    result.drift.push("jsaddons-missing");
    result.warnings = projectPublicWarnings(warnings);
    return result;
  }
  try {
    assertRealDirectory(layout.jsaddons, "jsaddons");
  } catch (error) {
    result.drift.push("jsaddons-path-error");
    warnings.push(error instanceof Error ? error.message : String(error));
    result.warnings = projectPublicWarnings(warnings);
    return result;
  }

  // publish
  let publishOk = false;
  if (fs.existsSync(layout.publishXml)) {
    try {
      assertRealFile(layout.publishXml, "publish.xml");
      const parsed = parseJspluginsDocument(fs.readFileSync(layout.publishXml, "utf8"));
      warnings.push(...parsed.warnings);
      const own = parsed.plugins.filter((p) => p.attrs.name === WPS_ADDON_NAME);
      if (own.length === 0) drift.push("publish-entry-missing");
      else if (own.length > 1) drift.push("publish-entry-duplicate");
      else if (!ownPluginMatchesContract(own[0].attrs)) drift.push("publish-entry-attrs");
      else publishOk = true;
    } catch (error) {
      drift.push("publish-parse-error");
      warnings.push(error instanceof Error ? error.message : String(error));
    }
  } else {
    drift.push("publish-missing");
  }

  // addon tree
  let addonPresent = false;
  /** @type {Map<string,string>|null} */
  let actualHashes = null;
  if (fs.existsSync(layout.addonDir)) {
    try {
      assertRealDirectory(layout.addonDir, "addonDir");
      addonPresent = true;
      actualHashes = hashAddonTree(layout.addonDir);
      for (const rel of [
        `${WPS_ADDON_DIRECTORY}/index.html`,
        `${WPS_ADDON_DIRECTORY}/manifest.xml`,
        `${WPS_ADDON_DIRECTORY}/ribbon.xml`,
        `${WPS_ADDON_DIRECTORY}/${WPS_ENTRY_SCRIPT}`,
      ]) {
        if (!actualHashes.has(rel)) drift.push(`addon-missing:${rel}`);
      }
    } catch (error) {
      drift.push("addon-path-error");
      warnings.push(error instanceof Error ? error.message : String(error));
      addonPresent = fs.existsSync(layout.addonDir);
    }
  } else {
    drift.push("addon-missing");
  }

  // state
  const stateRead = readStateFile(layout.stateFile);
  if (!stateRead.present) {
    drift.push("state-missing");
  } else if (stateRead.invalid) {
    drift.push(`state-invalid:${stateRead.reason || "unknown"}`);
  } else {
    const state = stateRead.state;
    result.restartRequired = state.restartRequired === true;
    result.packageVersion = state.packageVersion;
    result.gitSha = state.gitSha;

    if (!actualHashes) {
      drift.push("state-hash-unverified");
    } else {
      const expectedKeys = Object.keys(state.fileHashes).sort();
      const actualKeys = [...actualHashes.keys()].sort();
      const expectedSet = new Set(expectedKeys);
      const actualSet = new Set(actualKeys);
      for (const k of expectedKeys) {
        if (!actualSet.has(k)) drift.push(`hash-missing:${k}`);
        else if (actualHashes.get(k) !== state.fileHashes[k]) {
          drift.push(`hash-mismatch:${k}`);
        }
      }
      for (const k of actualKeys) {
        if (!expectedSet.has(k)) drift.push(`hash-extra:${k}`);
      }
    }
  }

  const noDrift = drift.length === 0;
  result.installed = Boolean(publishOk && addonPresent && stateRead.present && !stateRead.invalid);
  result.current = Boolean(result.installed && noDrift && publishOk);
  if (result.current) {
    result.message = "Installation looks current";
  } else if (result.installed || addonPresent || publishOk || stateRead.present) {
    result.message = "Installation present with drift";
    // keep installed true only if core pieces exist without claiming current
    if (!result.installed) {
      result.installed = Boolean(publishOk || addonPresent);
    }
  } else {
    result.message = "Not installed";
  }
  // Legacy kebab-case dir is never "current"; sanitized warning only.
  try {
    const legacy = inspectLegacyOwnAddon(layout);
    if (legacy.present) {
      warnings.push(
        legacy.verified
          ? "legacy own directory still present (run install to migrate)"
          : "legacy own directory present without verified state; not treated as current",
      );
      result.legacyOwnAddonPresent = true;
      result.legacyOwnAddonVerified = legacy.verified === true;
    } else {
      result.legacyOwnAddonPresent = false;
      result.legacyOwnAddonVerified = false;
    }
  } catch {
    result.legacyOwnAddonPresent = false;
    result.legacyOwnAddonVerified = false;
  }
  result.warnings = projectPublicWarnings(warnings);
  return result;
}
