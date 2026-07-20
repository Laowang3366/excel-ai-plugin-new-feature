/**
 * Shared WPS JSA install planning (read-only AppData inspection).
 * Used by both --dry-run and real install to avoid behavior drift.
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { createWpsPackage as defaultCreateWpsPackage } from "./package-wps-jsa.mjs";
import {
  hashMapToObject,
  packageDigest,
  validateWpsPackageDir,
} from "./wpsJsaInstallValidate.mjs";
import {
  emptyPublish,
  ownPluginMatchesContract,
  parseJspluginsDocument,
  upsertOwnPlugin,
} from "./wpsJsaInstallPublish.mjs";
import {
  assertAncestryReal,
  listActiveTempNames,
  lstatIfPresent,
  preflightExistingSurface,
  preflightOwnPrefixedEntries,
  resolveAppDataRoot,
  resolveJsaddonsLayout,
} from "./wpsJsaInstallPaths.mjs";
import { WPS_ADDON_DIRECTORY, WPS_ADDON_NAME } from "./wpsJsaPackage.mjs";
import {
  projectPublicPluginNames,
  projectPublicWarnings,
} from "./wpsJsaInstallPublicNames.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const defaultRoot = path.resolve(__dirname, "..");

/**
 * Resolve package directory and validate (may build dist if no packageDir).
 * @param {{
 *   packageDir?: string|null,
 *   gitSha?: string|null,
 *   rootDir?: string,
 *   skipBuild?: boolean,
 *   afterValidate?: (packageDir: string) => void,
 *   createWpsPackage?: Function,
 * }} opts
 */
export function resolveAndValidatePackage(opts = {}) {
  const createPkg = opts.createWpsPackage || defaultCreateWpsPackage;
  let packageDir = opts.packageDir ? path.resolve(opts.packageDir) : null;
  let built = false;

  if (!packageDir) {
    const summary = createPkg({
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
  validated = validateWpsPackageDir(packageDir);
  return { packageDir, built, validated };
}

/**
 * Read-only plan for installing into appData/jsaddons.
 * Never mutates appData (no mkdir/write/rename under appData).
 *
 * @param {{
 *   packageDir?: string|null,
 *   appData?: string|null,
 *   gitSha?: string|null,
 *   rootDir?: string,
 *   skipBuild?: boolean,
 *   platform?: string,
 *   env?: NodeJS.ProcessEnv,
 *   afterValidate?: (packageDir: string) => void,
 *   createWpsPackage?: Function,
 *   packageResolution?: { packageDir: string, built: boolean, validated: object },
 * }} opts
 */
export function planWpsJsaInstall(opts = {}) {
  const warnings = [];
  const resolved =
    opts.packageResolution ||
    resolveAndValidatePackage({
      packageDir: opts.packageDir,
      gitSha: opts.gitSha,
      rootDir: opts.rootDir,
      skipBuild: opts.skipBuild,
      afterValidate: opts.afterValidate,
      createWpsPackage: opts.createWpsPackage,
    });
  const { packageDir, built, validated } = resolved;

  const appData = resolveAppDataRoot({
    appData: opts.appData,
    platform: opts.platform,
    env: opts.env,
  });
  const layout = resolveJsaddonsLayout(appData);

  // Read-only ancestry preflight on every branch (no mkdir).
  // Fully absent appData is allowed; existing appData/kingsoft/wps ancestors
  // that are symlink/junction/non-directory fail closed so dry-run matches install.
  assertAncestryReal(layout.jsaddons, layout.appData);

  const jsaddonsSt = lstatIfPresent(layout.jsaddons);
  let wouldCreateJsaddons = false;
  /** @type {string[]} */
  let activeTemps = [];
  /** @type {string[]} */
  let preservedPluginNames = [];
  let existingOwnEntry = false;
  let wouldCreatePublish = false;
  let wouldUpdatePublish = true;
  let wouldReplaceAddon = false;
  let currentPublishBytes = null;
  let publishExisted = false;

  if (!jsaddonsSt) {
    wouldCreateJsaddons = true;
    wouldCreatePublish = true;
    wouldUpdatePublish = true;
    wouldReplaceAddon = false;
    const merged = upsertOwnPlugin(emptyPublish());
    warnings.push(...merged.warnings);
    return buildPlanResult({
      packageDir,
      built,
      validated,
      layout,
      appData,
      warnings,
      mergedXml: merged.xml,
      wouldCreateJsaddons,
      wouldCreatePublish,
      wouldUpdatePublish,
      wouldReplaceAddon,
      existingOwnEntry,
      preservedPluginNames,
      activeTemps,
      currentPublishBytes,
      publishExisted,
    });
  }

  // jsaddons exists — read-only preflight (symlink/type fail closed)
  if (jsaddonsSt.isSymbolicLink() || !jsaddonsSt.isDirectory()) {
    throw new Error(`jsaddons must be a real directory: ${layout.jsaddons}`);
  }
  preflightExistingSurface(layout);
  preflightOwnPrefixedEntries(layout.jsaddons);
  activeTemps = listActiveTempNames(layout.jsaddons);
  if (activeTemps.length > 0) {
    warnings.push(
      `active install temps present (not removed by dry-run/plan): ${activeTemps.join(", ")}`,
    );
  }

  const addonSt = lstatIfPresent(layout.addonDir);
  wouldReplaceAddon = Boolean(addonSt);

  const pubSt = lstatIfPresent(layout.publishXml);
  let currentPublish = emptyPublish();
  if (!pubSt) {
    wouldCreatePublish = true;
    wouldUpdatePublish = true;
  } else {
    publishExisted = true;
    if (pubSt.isSymbolicLink() || !pubSt.isFile()) {
      throw new Error(`publish.xml must be a regular file: ${layout.publishXml}`);
    }
    currentPublishBytes = fs.readFileSync(layout.publishXml, "utf8");
    currentPublish = currentPublishBytes;
    const parsed = parseJspluginsDocument(currentPublish);
    warnings.push(...parsed.warnings);
    preservedPluginNames = projectPublicPluginNames(
      parsed.plugins.map((p) => p.attrs.name).filter((n) => n && n !== WPS_ADDON_NAME),
    );
    const own = parsed.plugins.filter((p) => p.attrs.name === WPS_ADDON_NAME);
    existingOwnEntry = own.length === 1;
    if (own.length > 1) {
      throw new Error("publish.xml contains duplicate WenggeExcelAiAddin entries");
    }
  }

  const merged = upsertOwnPlugin(currentPublish);
  warnings.push(...merged.warnings);
  if (publishExisted) {
    // Semantic update only when serialized document changes
    wouldUpdatePublish = merged.xml !== currentPublishBytes;
    wouldCreatePublish = false;
  } else {
    wouldCreatePublish = true;
    wouldUpdatePublish = true;
  }

  // state presence is informational; install always writes state
  const stateSt = lstatIfPresent(layout.stateFile);
  if (stateSt && (stateSt.isSymbolicLink() || !stateSt.isFile())) {
    throw new Error(`state must be a regular file: ${layout.stateFile}`);
  }

  return buildPlanResult({
    packageDir,
    built,
    validated,
    layout,
    appData,
    warnings,
    mergedXml: merged.xml,
    wouldCreateJsaddons,
    wouldCreatePublish,
    wouldUpdatePublish,
    wouldReplaceAddon,
    existingOwnEntry,
    preservedPluginNames,
    activeTemps,
    currentPublishBytes,
    publishExisted,
    stateExisted: Boolean(stateSt),
  });
}

function buildPlanResult(ctx) {
  const addonHashes = new Map(
    [...ctx.validated.hashes.entries()].filter(([k]) =>
      k.startsWith(`${WPS_ADDON_DIRECTORY}/`),
    ),
  );
  const digest = packageDigest(ctx.validated.hashes);
  return {
    packageDir: ctx.packageDir,
    builtPackage: ctx.built,
    validated: ctx.validated,
    layout: ctx.layout,
    appData: ctx.appData,
    warnings: projectPublicWarnings(ctx.warnings || []),
    mergedXml: ctx.mergedXml,
    wouldCreateJsaddons: ctx.wouldCreateJsaddons,
    wouldCreatePublish: ctx.wouldCreatePublish,
    wouldUpdatePublish: ctx.wouldUpdatePublish,
    wouldReplaceAddon: ctx.wouldReplaceAddon,
    wouldWriteState: true,
    existingOwnEntry: ctx.existingOwnEntry,
    preservedPluginNames: ctx.preservedPluginNames,
    activeTemps: ctx.activeTemps,
    packageVersion: ctx.validated.buildInfo.packageVersion,
    gitSha: ctx.validated.buildInfo.gitSha,
    packageDigest: digest,
    addonFileCount: addonHashes.size,
    addonFileHashes: hashMapToObject(addonHashes),
    currentPublishBytes: ctx.currentPublishBytes,
    publishExisted: ctx.publishExisted,
    stateExisted: ctx.stateExisted === true,
  };
}

/**
 * Format public dry-run JSON (no secrets).
 */
export function formatDryRunResult(plan) {
  return {
    ok: true,
    action: "install",
    dryRun: true,
    wouldInstall: true,
    appData: plan.appData,
    jsaddons: plan.layout.jsaddons,
    addonDir: plan.layout.addonDir,
    publishXml: plan.layout.publishXml,
    stateFile: plan.layout.stateFile,
    packageDir: plan.packageDir,
    builtPackage: plan.builtPackage,
    packageVersion: plan.packageVersion,
    gitSha: plan.gitSha,
    packageDigest: plan.packageDigest,
    addonFileCount: plan.addonFileCount,
    wouldCreateJsaddons: plan.wouldCreateJsaddons,
    wouldReplaceAddon: plan.wouldReplaceAddon,
    wouldCreatePublish: plan.wouldCreatePublish,
    wouldUpdatePublish: plan.wouldUpdatePublish,
    wouldWriteState: true,
    existingOwnEntry: plan.existingOwnEntry,
    preservedPluginNames: projectPublicPluginNames(plan.preservedPluginNames || []),
    activeTemps: plan.activeTemps,
    warnings: projectPublicWarnings(plan.warnings || []),
    restartRequired: true,
    message:
      "Dry run only; no AppData/jsaddons files changed. Install may still rewrite publish.xml atomically and always writes state even when wouldUpdatePublish=false (entry attrs already match). Fully restart WPS after a real install. This tool does not start or stop WPS.",
  };
}
