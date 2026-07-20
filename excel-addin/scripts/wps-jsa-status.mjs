#!/usr/bin/env node
/** Report WPS JSA install status (honest drift; does not start WPS). */
import path from "node:path";
import { fileURLToPath } from "node:url";
import { statusWpsJsa } from "./wpsJsaInstallCore.mjs";
import { parseWpsInstallCliArgs } from "./wpsJsaInstallCliArgs.mjs";

function usage(code = 1) {
  console.error(`Usage:
  npm run wps:status -- [--app-data <dir>]
`);
  process.exit(code);
}

function main() {
  let args;
  try {
    args = parseWpsInstallCliArgs(process.argv.slice(2), {});
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
