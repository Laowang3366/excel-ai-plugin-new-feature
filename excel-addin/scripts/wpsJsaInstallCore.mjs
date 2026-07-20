/**
 * WPS JSA install / status / uninstall orchestration (install-time Node CLI only).
 */
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
  ownPluginMatchesContract,
  parseJspluginsDocument,
  removeOwnPlugin,
  upsertOwnPlugin,
  writePublishXmlAtomic,
} from "./wpsJsaInstallPublish.mjs";
import {
  assertInside,
  assertRealDirectory,
  assertRealFile,
  ensureJsaddonsDir,
  prevAddonDir,
  resolveAppDataRoot,
  resolveJsaddonsLayout,
  safeRenameInside,
  safeRmInside,
  stagingDir,
} from "./wpsJsaInstallPaths.mjs";
import { WPS_ADDON_DIRECTORY, WPS_ADDON_NAME, WPS_ENTRY_SCRIPT } from "./wpsJsaPackage.mjs";
import { listFilesRecursiveStrict } from "./packageProdCore.mjs";
import { createHash } from "node:crypto";

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

function verifyStagedAddon(stageAddonDir, packageValidation) {
  // staged addon content should match hashes under wengge-excel-ai-addin/*
  const prefix = `${WPS_ADDON_DIRECTORY}/`;
  for (const [rel, expected] of packageValidation.hashes) {
    if (!rel.startsWith(prefix)) continue;
    const sub = rel.slice(prefix.length);
    const abs = path.join(stageAddonDir, ...sub.split("/"));
    assertRealFile(abs, rel);
    const actual = createHash("sha256").update(fs.readFileSync(abs)).digest("hex");
    if (actual !== expected) {
      throw new Error(`staged hash mismatch for ${rel}`);
    }
  }
}

function writeStateAtomic(jsaddons, statePath, state) {
  assertInside(jsaddons, statePath, "state");
  const tmp = path.join(jsaddons, `${path.basename(statePath)}.tmp.${Date.now()}`);
  assertInside(jsaddons, tmp, "state tmp");
  fs.writeFileSync(tmp, `${JSON.stringify(state, null, 2)}\n`, "utf8");
  const fd = fs.openSync(tmp, "r+");
  try {
    fs.fsyncSync(fd);
  } finally {
    fs.closeSync(fd);
  }
  fs.renameSync(tmp, statePath);
}

function readState(statePath) {
  if (!fs.existsSync(statePath)) return null;
  assertRealFile(statePath, "state");
  try {
    return JSON.parse(fs.readFileSync(statePath, "utf8"));
  } catch {
    return { _invalid: true };
  }
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

  // Validate fully before any appData mutation
  const validated = validateWpsPackageDir(packageDir);

  const appData = resolveAppDataRoot({
    appData: opts.appData,
    platform: opts.platform,
    env: opts.env,
  });
  const layout = resolveJsaddonsLayout(appData);
  ensureJsaddonsDir(layout.jsaddons, layout.appData);

  const stageRoot = stagingDir(layout.jsaddons);
  assertInside(layout.jsaddons, stageRoot, "staging");
  const stageAddon = path.join(stageRoot, WPS_ADDON_DIRECTORY);
  let prevDir = null;
  let publishBackup = null;

  try {
    fs.mkdirSync(stageRoot);
    copyDirReal(validated.addonDir, stageAddon, layout.jsaddons);
    verifyStagedAddon(stageAddon, validated);

    // Prepare publish content before swap
    let currentPublish = emptyPublish();
    if (fs.existsSync(layout.publishXml)) {
      assertRealFile(layout.publishXml, "publish.xml");
      currentPublish = fs.readFileSync(layout.publishXml, "utf8");
    }
    const merged = upsertOwnPlugin(currentPublish);
    warnings.push(...merged.warnings);

    // Swap addon directory
    if (fs.existsSync(layout.addonDir)) {
      assertRealDirectory(layout.addonDir, "existing addon");
      prevDir = prevAddonDir(layout.jsaddons);
      safeRenameInside(layout.jsaddons, layout.addonDir, prevDir);
    }
    safeRenameInside(layout.jsaddons, stageAddon, layout.addonDir);
    // stageRoot may still exist empty-ish
    try {
      safeRmInside(layout.jsaddons, stageRoot);
    } catch {
      /* best effort */
    }

    try {
      const written = writePublishXmlAtomic(layout.jsaddons, layout.publishXml, merged.xml);
      publishBackup = written.backedUp;
    } catch (error) {
      // rollback addon
      if (fs.existsSync(layout.addonDir)) {
        safeRmInside(layout.jsaddons, layout.addonDir);
      }
      if (prevDir && fs.existsSync(prevDir)) {
        safeRenameInside(layout.jsaddons, prevDir, layout.addonDir);
      }
      throw error;
    }

    // success: remove previous addon backup
    if (prevDir && fs.existsSync(prevDir)) {
      safeRmInside(layout.jsaddons, prevDir);
      prevDir = null;
    }

    const state = {
      schemaVersion: 1,
      addonName: WPS_ADDON_NAME,
      addonDirectory: WPS_ADDON_DIRECTORY,
      installedAt: new Date().toISOString(),
      packageVersion: validated.buildInfo.packageVersion,
      gitSha: validated.buildInfo.gitSha,
      publishUrl: validated.buildInfo.publishUrl,
      packageDigest: packageDigest(validated.hashes),
      fileHashes: hashMapToObject(
        new Map(
          [...validated.hashes.entries()].filter(([k]) =>
            k.startsWith(`${WPS_ADDON_DIRECTORY}/`),
          ),
        ),
      ),
      restartRequired: true,
      builtPackage: built,
    };
    writeStateAtomic(layout.jsaddons, layout.stateFile, state);

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
    };
  } catch (error) {
    // cleanup staging
    try {
      if (fs.existsSync(stageRoot)) safeRmInside(layout.jsaddons, stageRoot);
    } catch {
      /* ignore */
    }
    // if prev exists and addon missing, restore
    try {
      if (prevDir && fs.existsSync(prevDir) && !fs.existsSync(layout.addonDir)) {
        safeRenameInside(layout.jsaddons, prevDir, layout.addonDir);
      }
    } catch {
      /* ignore */
    }
    throw error;
  }
}

export function statusWpsJsa(opts = {}) {
  const warnings = [];
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
    drift: [],
    warnings,
    restartRequired: null,
    publishXml: layout.publishXml,
    addonDir: layout.addonDir,
    stateFile: layout.stateFile,
  };

  if (!fs.existsSync(layout.jsaddons)) {
    result.message = "jsaddons directory not found";
    return result;
  }
  assertRealDirectory(layout.jsaddons, "jsaddons");

  let publishOk = false;
  let ownAttrs = null;
  if (fs.existsSync(layout.publishXml)) {
    assertRealFile(layout.publishXml, "publish.xml");
    try {
      const parsed = parseJspluginsDocument(fs.readFileSync(layout.publishXml, "utf8"));
      warnings.push(...parsed.warnings);
      const own = parsed.plugins.filter((p) => p.attrs.name === WPS_ADDON_NAME);
      if (own.length === 1) {
        ownAttrs = own[0].attrs;
        publishOk = ownPluginMatchesContract(ownAttrs);
        if (!publishOk) result.drift.push("publish-entry-attrs");
      } else if (own.length === 0) {
        result.drift.push("publish-entry-missing");
      } else {
        result.drift.push("publish-entry-duplicate");
      }
    } catch (error) {
      result.drift.push("publish-parse-error");
      warnings.push(error instanceof Error ? error.message : String(error));
    }
  } else {
    result.drift.push("publish-missing");
  }

  const addonPresent = fs.existsSync(layout.addonDir);
  if (addonPresent) {
    try {
      assertRealDirectory(layout.addonDir, "addonDir");
      for (const rel of [
        "index.html",
        "manifest.xml",
        "ribbon.xml",
        WPS_ENTRY_SCRIPT,
      ]) {
        const abs = path.join(layout.addonDir, rel);
        if (!fs.existsSync(abs)) {
          result.drift.push(`addon-missing:${rel}`);
        } else {
          assertRealFile(abs, rel);
        }
      }
    } catch (error) {
      result.drift.push("addon-path-error");
      warnings.push(error instanceof Error ? error.message : String(error));
    }
  } else {
    result.drift.push("addon-missing");
  }

  const state = readState(layout.stateFile);
  if (!state) {
    result.drift.push("state-missing");
  } else if (state._invalid) {
    result.drift.push("state-invalid");
  } else {
    result.restartRequired = state.restartRequired === true;
    result.packageVersion = state.packageVersion;
    result.gitSha = state.gitSha;
    // hash drift vs state
    if (addonPresent && state.fileHashes && typeof state.fileHashes === "object") {
      for (const [rel, expected] of Object.entries(state.fileHashes)) {
        if (!rel.startsWith(`${WPS_ADDON_DIRECTORY}/`)) continue;
        const sub = rel.slice(WPS_ADDON_DIRECTORY.length + 1);
        const abs = path.join(layout.addonDir, ...sub.split("/"));
        if (!fs.existsSync(abs)) {
          result.drift.push(`hash-missing:${rel}`);
          continue;
        }
        try {
          const actual = createHash("sha256").update(fs.readFileSync(abs)).digest("hex");
          if (actual !== expected) result.drift.push(`hash-mismatch:${rel}`);
        } catch {
          result.drift.push(`hash-read-error:${rel}`);
        }
      }
    }
  }

  result.installed = Boolean(
    publishOk && addonPresent && state && !state._invalid && result.drift.length === 0,
  );
  // softer installed: present pieces even with drift
  const present =
    (ownAttrs != null || addonPresent || (state && !state._invalid));
  if (present && !result.installed) {
    result.installed = ownAttrs != null || addonPresent;
    result.current = false;
  } else {
    result.current = result.installed;
  }
  if (result.installed && result.drift.length === 0) {
    result.current = true;
    result.message = "Installation looks current";
  } else if (result.installed) {
    result.message = "Installation present with drift";
  } else {
    result.message = "Not installed";
  }
  return result;
}

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

  // publish: remove own entry if file exists
  if (fs.existsSync(layout.publishXml)) {
    assertRealFile(layout.publishXml, "publish.xml");
    const current = fs.readFileSync(layout.publishXml, "utf8");
    try {
      const removed = removeOwnPlugin(current);
      warnings.push(...removed.warnings);
      if (removed.removed || current.trim() !== removed.xml.trim()) {
        writePublishXmlAtomic(layout.jsaddons, layout.publishXml, removed.xml);
      }
    } catch (error) {
      // If parse fails, do not destroy foreign publish; surface error
      throw new Error(
        `cannot safely edit publish.xml during uninstall: ${
          error instanceof Error ? error.message : String(error)
        }`,
      );
    }
  }

  if (fs.existsSync(layout.addonDir)) {
    assertRealDirectory(layout.addonDir, "addonDir");
    // ensure name is exact
    if (path.basename(layout.addonDir) !== WPS_ADDON_DIRECTORY) {
      throw new Error("refusing to remove unexpected addon directory name");
    }
    safeRmInside(layout.jsaddons, layout.addonDir);
  }

  if (fs.existsSync(layout.stateFile)) {
    assertRealFile(layout.stateFile, "state");
    if (path.basename(layout.stateFile) !== path.basename(layout.stateFile)) {
      throw new Error("state basename mismatch");
    }
    fs.unlinkSync(layout.stateFile);
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
}

// re-export helpers for tests
export { validateWpsPackageDir, parseJspluginsDocument, upsertOwnPlugin, removeOwnPlugin };
