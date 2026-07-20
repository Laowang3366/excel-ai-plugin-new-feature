#!/usr/bin/env node
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { validateWpsSourceBundle } from "./wpsJsaPackage.mjs";

const rootDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const sourceDir = path.join(rootDir, "manifest/wps-jsa");
const read = (name) => fs.readFileSync(path.join(sourceDir, name), "utf8");

const validation = validateWpsSourceBundle({
  manifestXml: read("manifest.xml"),
  ribbonXml: read("ribbon.xml"),
  entryScript: read("wps-entry.js"),
  publishXml: read("publish.xml"),
});

if (!validation.ok) {
  console.error("WPS JSA package source validation failed:");
  for (const error of validation.errors) console.error(`  - ${error}`);
  process.exit(1);
}

console.log("WPS JSA manifest, Ribbon, entry, and publish.xml check OK.");
