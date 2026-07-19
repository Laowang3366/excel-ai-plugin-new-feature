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
const checkedIn = path.join(rootDir, "manifest/office-excel-manifest.xml");
const templatePath = path.join(
  rootDir,
  "manifest/templates/office-excel-manifest.template.xml",
);

const template = fs.readFileSync(templatePath, "utf8");
const expected = renderOfficeManifest({
  mode: "dev",
  baseUrl: DEFAULT_DEV_BASE_URL,
  version: DEFAULT_VERSION,
  template,
});
const actual = fs.readFileSync(checkedIn, "utf8");

const validation = validateOfficeManifest(actual, { mode: "dev" });
let failed = false;
if (!validation.ok) {
  failed = true;
  console.error("Checked-in manifest validation failed:");
  for (const e of validation.errors) console.error(`  - ${e}`);
}

const norm = (s) => s.replace(/\r\n/g, "\n").trimEnd() + "\n";
if (norm(actual) !== norm(expected)) {
  failed = true;
  console.error(
    "Checked-in manifest/office-excel-manifest.xml drifts from template render (dev defaults).",
  );
  console.error("Run: npm run manifest:dev");
}

if (failed) process.exit(1);
console.log("Office manifest check OK (dev defaults + validation).");
