#!/usr/bin/env node
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import {
  DEFAULT_DEV_BASE_URL,
  DEFAULT_VERSION,
  renderOfficeManifest,
  validateOfficeManifest,
} from "./officeManifest.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.resolve(__dirname, "..");

function usage(code = 1) {
  console.error(`Usage:
  node scripts/render-office-manifest.mjs --mode dev|prod [--base-url URL] [--version X.Y.Z.W] [--out PATH]

Defaults:
  dev  base-url=${DEFAULT_DEV_BASE_URL}
  prod base-url required (HTTPS, non-localhost)
  out  manifest/office-excel-manifest.xml (dev) or stdout if --out -
`);
  process.exit(code);
}

function parseArgs(argv) {
  const out = {
    mode: null,
    baseUrl: null,
    version: DEFAULT_VERSION,
    out: null,
  };
  for (let i = 0; i < argv.length; i += 1) {
    const a = argv[i];
    if (a === "--mode") out.mode = argv[++i];
    else if (a === "--base-url") out.baseUrl = argv[++i];
    else if (a === "--version") out.version = argv[++i];
    else if (a === "--out") out.out = argv[++i];
    else if (a === "-h" || a === "--help") usage(0);
    else {
      console.error(`Unknown arg: ${a}`);
      usage(1);
    }
  }
  return out;
}

const args = parseArgs(process.argv.slice(2));
if (args.mode !== "dev" && args.mode !== "prod") usage(1);
if (args.mode === "prod" && !args.baseUrl) {
  console.error("prod requires --base-url https://...");
  process.exit(1);
}

const templatePath = path.join(
  rootDir,
  "manifest/templates/office-excel-manifest.template.xml",
);
const template = fs.readFileSync(templatePath, "utf8");
let xml;
try {
  xml = renderOfficeManifest({
    mode: args.mode,
    baseUrl: args.baseUrl ?? DEFAULT_DEV_BASE_URL,
    version: args.version,
    template,
  });
} catch (err) {
  console.error(String(err?.message || err));
  process.exit(1);
}

const validation = validateOfficeManifest(xml, { mode: args.mode });
if (!validation.ok) {
  console.error("Manifest validation failed:");
  for (const e of validation.errors) console.error(`  - ${e}`);
  process.exit(1);
}

const defaultOut =
  args.mode === "dev"
    ? path.join(rootDir, "manifest/office-excel-manifest.xml")
    : null;
const outPath = args.out ?? defaultOut;
if (!outPath || outPath === "-") {
  process.stdout.write(xml);
} else {
  const abs = path.isAbsolute(outPath) ? outPath : path.resolve(process.cwd(), outPath);
  fs.mkdirSync(path.dirname(abs), { recursive: true });
  fs.writeFileSync(abs, xml, "utf8");
  console.log(`Wrote ${abs}`);
}
