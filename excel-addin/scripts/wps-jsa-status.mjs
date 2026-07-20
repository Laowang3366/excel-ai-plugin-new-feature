#!/usr/bin/env node
/** Report WPS JSA install status (honest drift; does not start WPS). */
import path from "node:path";
import { fileURLToPath } from "node:url";
import { statusWpsJsa } from "./wpsJsaInstallCore.mjs";

function usage(code = 1) {
  console.error(`Usage:
  npm run wps:status -- [--app-data <dir>]
`);
  process.exit(code);
}

function parseArgs(argv) {
  const out = { appData: null, help: false };
  for (let i = 0; i < argv.length; i += 1) {
    const a = argv[i];
    if (a === "--app-data") out.appData = argv[++i] ?? null;
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
    const result = statusWpsJsa({ appData: args.appData });
    process.stdout.write(`${JSON.stringify(result, null, 2)}\n`);
    if (!result.current) process.exitCode = 2;
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    process.exitCode = 1;
  }
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main();
}
