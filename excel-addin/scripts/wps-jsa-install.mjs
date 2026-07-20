#!/usr/bin/env node
/** Install WPS JSA package into local jsaddons (does not start/stop WPS). */
import path from "node:path";
import { fileURLToPath } from "node:url";
import { installWpsJsa } from "./wpsJsaInstallCore.mjs";

function usage(code = 1) {
  console.error(`Usage:
  npm run wps:install -- [--git-sha <sha>] [--package-dir <dir>] [--app-data <dir>]

  Default: build package via package:wps flow into dist/, then install.
  --package-dir: install an existing package directory (no rebuild).
  --app-data: override AppData root (required on non-Windows / tests).

Never starts, stops, or attaches to WPS; always reports restartRequired.
`);
  process.exit(code);
}

function parseArgs(argv) {
  const out = {
    gitSha: null,
    packageDir: null,
    appData: null,
    help: false,
  };
  for (let i = 0; i < argv.length; i += 1) {
    const a = argv[i];
    if (a === "--git-sha") out.gitSha = argv[++i] ?? null;
    else if (a === "--package-dir") out.packageDir = argv[++i] ?? null;
    else if (a === "--app-data") out.appData = argv[++i] ?? null;
    else if (a === "-h" || a === "--help") out.help = true;
    else throw new Error(`Unknown arg: ${a}`);
  }
  return out;
}

function main() {
  let args;
  try {
    args = parseArgs(process.argv.slice(2));
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
