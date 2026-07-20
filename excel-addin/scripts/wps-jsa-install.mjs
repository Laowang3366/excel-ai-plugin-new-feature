#!/usr/bin/env node
/** Install WPS JSA package into local jsaddons (does not start/stop WPS). */
import path from "node:path";
import { fileURLToPath } from "node:url";
import { installWpsJsa } from "./wpsJsaInstallCore.mjs";
import { parseWpsInstallCliArgs } from "./wpsJsaInstallCliArgs.mjs";

function usage(code = 1) {
  console.error(`Usage:
  npm run wps:install -- [--dry-run] [--git-sha <sha>] [--package-dir <dir>] [--app-data <dir>]

  Default: build package via package:wps flow into dist/, then install.
  --package-dir: install an existing package directory (no rebuild; do not pass --git-sha).
  --dry-run: plan-only; guarantees zero AppData/jsaddons writes.
             If --package-dir is omitted, may still rebuild project dist/.
             For real AppData zero-write preview: package:wps first, then
             --package-dir ./dist --dry-run --app-data <AppData>.
  --app-data: override AppData root (required on non-Windows / tests).

Never starts, stops, or attaches to WPS; always reports restartRequired.
`);
  process.exit(code);
}

function main() {
  let args;
  try {
    args = parseWpsInstallCliArgs(process.argv.slice(2), {
      allowGitSha: true,
      allowPackageDir: true,
      allowDryRun: true,
    });
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    usage(1);
  }
  if (args.help) usage(0);
  try {
    const result = installWpsJsa({
      gitSha: args.gitSha,
      packageDir: args.packageDir,
      appData: args.appData,
      dryRun: args.dryRun === true,
    });
    process.stdout.write(`${JSON.stringify(result, null, 2)}\n`);
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    process.exitCode = 1;
  }
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main();
}

export { parseWpsInstallCliArgs };
