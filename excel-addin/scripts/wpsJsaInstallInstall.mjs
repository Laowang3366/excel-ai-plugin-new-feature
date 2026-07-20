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
} from "./wpsJsaInstallState.mjs";
import {
  assertInside,
  assertRealDirectory,
  assertRealFile,
  ensureJsaddonsDir,
  preflightExistingSurface,
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

function verifyStagedAddon(stageAddonDir, packageValidation) {
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

function snapshotBytesIfExists(filePath, label) {
  if (!fs.existsSync(filePath)) {
    return { existed: false, bytes: null };
  }
  assertRealFile(filePath, label);
  return { existed: true, bytes: fs.readFileSync(filePath, "utf8") };
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
 *   failAfter?: 'addon-swap'|'publish-write'|'state-write'|null,
 *   failpoints?: Record<string, Function>,
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

  // Full validation before any appData mutation
  const validated = validateWpsPackageDir(packageDir);

  const appData = resolveAppDataRoot({
    appData: opts.appData,
    platform: opts.platform,
    env: opts.env,
  });
  const layout = resolveJsaddonsLayout(appData);
  ensureJsaddonsDir(layout.jsaddons, layout.appData);
  preflightExistingSurface(layout);

  // Snapshot old surface for true rollback
  const oldPublish = snapshotBytesIfExists(layout.publishXml, "publish.xml");
  const oldState = snapshotBytesIfExists(layout.stateFile, "state");
  const hadOldAddon = fs.existsSync(layout.addonDir);
  if (hadOldAddon) {
    assertRealDirectory(layout.addonDir, "existing addon");
  }

  // Prepare merged publish before mutating install dirs
  let currentPublish = emptyPublish();
  if (oldPublish.existed) currentPublish = oldPublish.bytes;
  const merged = upsertOwnPlugin(currentPublish);
  warnings.push(...merged.warnings);

  const stageRoot = stagingDir(layout.jsaddons);
  const stageAddon = path.join(stageRoot, WPS_ADDON_DIRECTORY);
  let prevDir = null;
  let swapped = false;
  let publishCommitted = false;
  let stateCommitted = false;
  let publishMeta = { previousBytes: oldPublish.bytes, previousExisted: oldPublish.existed };
  let stateMeta = { previousBytes: oldState.bytes, previousExisted: oldState.existed };

  const failAfter = opts.failAfter || null;
  const fp = opts.failpoints || {};

  function runFail(name) {
    if (failAfter === name) throw new Error(`failpoint:${name}`);
    if (typeof fp[name] === "function") fp[name]();
  }

  function rollback() {
    // Restore addon
    try {
      if (swapped) {
        if (fs.existsSync(layout.addonDir)) {
          safeRmInside(layout.jsaddons, layout.addonDir);
        }
        if (prevDir && fs.existsSync(prevDir)) {
          safeRenameInside(layout.jsaddons, prevDir, layout.addonDir);
          prevDir = null;
        }
      } else if (prevDir && fs.existsSync(prevDir) && !fs.existsSync(layout.addonDir)) {
        safeRenameInside(layout.jsaddons, prevDir, layout.addonDir);
        prevDir = null;
      }
    } catch {
      /* best effort */
    }
    // Restore publish
    try {
      if (publishCommitted || fs.existsSync(layout.publishXml)) {
        restorePublishBytes(
          layout.jsaddons,
          layout.publishXml,
          publishMeta.previousBytes,
          publishMeta.previousExisted,
        );
      }
    } catch {
      /* best effort */
    }
    // Restore state
    try {
      restoreStateBytes(
        layout.jsaddons,
        layout.stateFile,
        stateMeta.previousBytes,
        stateMeta.previousExisted,
      );
    } catch {
      /* best effort */
    }
    // Cleanup staging/prev leftovers
    try {
      if (fs.existsSync(stageRoot)) safeRmInside(layout.jsaddons, stageRoot);
    } catch {
      /* ignore */
    }
    try {
      if (prevDir && fs.existsSync(prevDir)) safeRmInside(layout.jsaddons, prevDir);
    } catch {
      /* ignore */
    }
  }

  try {
    copyDirReal(validated.addonDir, stageAddon, layout.jsaddons);
    verifyStagedAddon(stageAddon, validated);

    // Move old addon aside (kept until publish+state succeed)
    if (hadOldAddon) {
      prevDir = prevAddonDir(layout.jsaddons);
      // prevAddonDir creates an empty exclusive dir — remove it so rename can use the path
      fs.rmdirSync(prevDir);
      safeRenameInside(layout.jsaddons, layout.addonDir, prevDir);
    }

    // Place new addon
    safeRenameInside(layout.jsaddons, stageAddon, layout.addonDir);
    swapped = true;
    try {
      if (fs.existsSync(stageRoot)) safeRmInside(layout.jsaddons, stageRoot);
    } catch {
      /* best effort */
    }
    runFail("addon-swap");

    // Publish commit
    publishMeta = {
      ...writePublishXmlAtomic(layout.jsaddons, layout.publishXml, merged.xml, {
        failBeforeRename: () => runFail("publish-write"),
      }),
      previousBytes: oldPublish.bytes,
      previousExisted: oldPublish.existed,
    };
    publishCommitted = true;
    runFail("publish-write-after");

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

    stateMeta = {
      ...writeStateAtomic(layout.jsaddons, layout.stateFile, state, {
        failBeforeRename: () => runFail("state-write"),
      }),
      previousBytes: oldState.bytes,
      previousExisted: oldState.existed,
    };
    stateCommitted = true;
    runFail("state-write-after");

    // Success: cleanup previous addon; failure here is warning only
    if (prevDir && fs.existsSync(prevDir)) {
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
    };
  } catch (error) {
    rollback();
    throw error;
  }
}
