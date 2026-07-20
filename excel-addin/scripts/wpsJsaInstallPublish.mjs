/**
 * Restricted deterministic publish.xml tokenizer/parser for jsaddons.
 */
import fs from "node:fs";
import path from "node:path";
import { WPS_ADDON_NAME, WPS_PUBLISH_URL } from "./wpsJsaPackage.mjs";
import {
  assertInside,
  assertOwnPublishBackupSurface,
  exclusiveTempFile,
  lstatIfPresent,
  ownPublishBackupPath,
  rotateOwnPublishBackups,
  TMP_PREFIX,
} from "./wpsJsaInstallPaths.mjs";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8"?>';
// version=1.0 required; encoding optional, only utf-8 (any case)
const XML_DECL_RE =
  /^<\?xml\s+version\s*=\s*(["'])1\.0\1(?:\s+encoding\s*=\s*(["'])utf-8\2)?\s*\?>$/i;

export function renderOwnJsplugin() {
  return `  <jsplugin
    name="${WPS_ADDON_NAME}"
    type="et"
    url="${WPS_PUBLISH_URL}"
    debug=""
    enable="enable_dev" />`;
}

export function emptyPublish() {
  return `${XML_DECL}\n<jsplugins>\n</jsplugins>\n`;
}

function skipWs(text, i) {
  while (i < text.length && /[ \t\r\n]/.test(text[i])) i += 1;
  return i;
}

function parseXmlDeclaration(text, i) {
  i = skipWs(text, i);
  if (!text.startsWith("<?xml", i)) return i;
  const end = text.indexOf("?>", i);
  if (end < 0) throw new Error("unterminated XML declaration");
  const decl = text.slice(i, end + 2);
  if (!XML_DECL_RE.test(decl)) {
    throw new Error(
      "publish.xml XML declaration must be version=1.0 with optional encoding=utf-8 only",
    );
  }
  return end + 2;
}

export function parseAttributes(attrText) {
  const attrs = {};
  let i = 0;
  const s = attrText;
  while (true) {
    i = skipWs(s, i);
    if (i >= s.length) break;
    const nameMatch = /^([A-Za-z_][\w:.-]*)/.exec(s.slice(i));
    if (!nameMatch) throw new Error(`invalid attribute near: ${s.slice(i, i + 20)}`);
    const name = nameMatch[1];
    i += name.length;
    i = skipWs(s, i);
    if (s[i] !== "=") throw new Error(`attribute ${name} missing =`);
    i += 1;
    i = skipWs(s, i);
    const q = s[i];
    if (q !== '"' && q !== "'") throw new Error(`attribute ${name} value must be quoted`);
    i += 1;
    let value = "";
    while (i < s.length && s[i] !== q) {
      if (s[i] === "<" || s[i] === ">") {
        throw new Error(`unsafe character in attribute ${name}`);
      }
      value += s[i];
      i += 1;
    }
    if (i >= s.length) throw new Error(`unterminated attribute value for ${name}`);
    i += 1;
    if (Object.prototype.hasOwnProperty.call(attrs, name)) {
      throw new Error(`duplicate attribute: ${name}`);
    }
    attrs[name] = value;
  }
  return attrs;
}

export function parseJspluginsDocument(xml) {
  const warnings = [];
  const text = String(xml).replace(/\r\n/g, "\n");
  if (text.trim() === "") throw new Error("publish.xml is empty");
  if (/<!DOCTYPE/i.test(text) || /<!ENTITY/i.test(text)) {
    throw new Error("publish.xml must not contain DOCTYPE/ENTITY");
  }
  if (/<!--/.test(text) || /<!\[CDATA\[/i.test(text)) {
    throw new Error("publish.xml comments/CDATA are not supported");
  }
  let i = 0;
  i = parseXmlDeclaration(text, i);
  i = skipWs(text, i);
  if (!text.startsWith("<jsplugins", i)) {
    throw new Error("publish.xml must start with jsplugins root");
  }
  const openEnd = text.indexOf(">", i);
  if (openEnd < 0) throw new Error("unterminated jsplugins open tag");
  const openTag = text.slice(i, openEnd + 1);
  if (!/^<jsplugins\s*>$/i.test(openTag)) {
    throw new Error("publish.xml jsplugins root must not have attributes");
  }
  i = openEnd + 1;

  const plugins = [];
  while (true) {
    i = skipWs(text, i);
    if (text.startsWith("</jsplugins>", i)) {
      i += "</jsplugins>".length;
      break;
    }
    if (i >= text.length) throw new Error("publish.xml missing closing jsplugins");
    if (!text.startsWith("<jsplugin", i)) {
      throw new Error("publish.xml contains non-jsplugin markup");
    }
    let j = i + "<jsplugin".length;
    let inQuote = null;
    while (j < text.length) {
      const ch = text[j];
      if (inQuote) {
        if (ch === inQuote) inQuote = null;
        j += 1;
        continue;
      }
      if (ch === '"' || ch === "'") {
        inQuote = ch;
        j += 1;
        continue;
      }
      if (ch === "<") {
        throw new Error("nested markup inside jsplugin is forbidden");
      }
      if (ch === ">") {
        const tag = text.slice(i, j + 1);
        if (!/\/\s*>$/.test(tag)) {
          throw new Error("publish.xml jsplugin must be self-closing");
        }
        const inner = tag.replace(/^<jsplugin/i, "").replace(/\/\s*>$/, "");
        const attrs = parseAttributes(inner);
        if (!attrs.name) throw new Error("jsplugin missing name attribute");
        plugins.push({ raw: tag, attrs });
        j += 1;
        i = j;
        break;
      }
      j += 1;
    }
    if (j >= text.length) throw new Error("unterminated jsplugin tag");
  }

  i = skipWs(text, i);
  if (i !== text.length) {
    throw new Error("publish.xml has trailing content after jsplugins");
  }

  const names = plugins.map((p) => p.attrs.name);
  if (names.filter((n) => n === WPS_ADDON_NAME).length > 1) {
    throw new Error("publish.xml contains duplicate WenggeExcelAiAddin entries");
  }
  for (const n of names) {
    if (n && n !== WPS_ADDON_NAME && /excelaiwps|excel-ai-wps|wenggeexcel/i.test(n)) {
      warnings.push(`legacy or third-party plugin present: ${n}`);
    }
  }
  return { plugins, warnings };
}

function renderDocument(plugins) {
  const body = plugins
    .map((p) => {
      if (p.attrs.name === WPS_ADDON_NAME) return renderOwnJsplugin();
      return `  ${p.raw.trim()}`;
    })
    .join("\n");
  return `${XML_DECL}\n<jsplugins>\n${body ? `${body}\n` : ""}</jsplugins>\n`;
}

export function upsertOwnPlugin(xml) {
  const parsed = parseJspluginsDocument(xml);
  const warnings = [...parsed.warnings];
  const others = parsed.plugins.filter((p) => p.attrs.name !== WPS_ADDON_NAME);
  const next = [
    ...others,
    {
      raw: renderOwnJsplugin().trim(),
      attrs: {
        name: WPS_ADDON_NAME,
        type: "et",
        url: WPS_PUBLISH_URL,
        debug: "",
        enable: "enable_dev",
      },
    },
  ];
  return { xml: renderDocument(next), warnings, plugins: next };
}

export function removeOwnPlugin(xml) {
  const parsed = parseJspluginsDocument(xml);
  const warnings = [...parsed.warnings];
  const next = parsed.plugins.filter((p) => p.attrs.name !== WPS_ADDON_NAME);
  const removed = parsed.plugins.length !== next.length;
  return { xml: renderDocument(next), warnings, removed, plugins: next };
}

/** Own entry must be exactly the five contract attributes (order free). */
export function ownPluginMatchesContract(attrs) {
  if (!attrs || typeof attrs !== "object") return false;
  const keys = Object.keys(attrs).sort();
  const expected = ["debug", "enable", "name", "type", "url"];
  if (keys.length !== expected.length || keys.join(",") !== expected.join(",")) {
    return false;
  }
  return (
    attrs.name === WPS_ADDON_NAME &&
    attrs.type === "et" &&
    attrs.url === WPS_PUBLISH_URL &&
    attrs.debug === "" &&
    attrs.enable === "enable_dev"
  );
}

/**
 * Atomic publish write. Backup rotation is best-effort AFTER successful commit.
 */
export function writePublishXmlAtomic(jsaddons, publishPath, newXml, opts = {}) {
  assertInside(jsaddons, publishPath, "publish.xml");
  if (path.basename(publishPath) !== "publish.xml") {
    throw new Error("publish path basename must be publish.xml");
  }
  // pre-commit: read-only validate backup surface (no rotate yet)
  assertOwnPublishBackupSurface(jsaddons);

  let previousBytes = null;
  let previousExisted = false;
  let backedUp = null;

  const existing = lstatIfPresent(publishPath);
  if (existing) {
    if (existing.isSymbolicLink() || !existing.isFile()) {
      throw new Error(`publish.xml must be a regular file: ${publishPath}`);
    }
    previousExisted = true;
    previousBytes = fs.readFileSync(publishPath, "utf8");
    backedUp = ownPublishBackupPath(jsaddons);
    fs.writeFileSync(backedUp, previousBytes, "utf8");
    const bfd = fs.openSync(backedUp, "r+");
    try {
      fs.fsyncSync(bfd);
    } finally {
      fs.closeSync(bfd);
    }
  }

  const tmp = exclusiveTempFile(jsaddons, `${TMP_PREFIX}publish-`);
  try {
    if (typeof opts.failBeforeWrite === "function") opts.failBeforeWrite();
    fs.writeFileSync(tmp, newXml, "utf8");
    const fd = fs.openSync(tmp, "r+");
    try {
      fs.fsyncSync(fd);
    } finally {
      fs.closeSync(fd);
    }
    if (typeof opts.failBeforeRename === "function") opts.failBeforeRename();
    fs.renameSync(tmp, publishPath);
    if (typeof opts.failAfterCommit === "function") opts.failAfterCommit();
  } catch (error) {
    try {
      if (lstatIfPresent(tmp) && !lstatIfPresent(tmp).isSymbolicLink()) {
        fs.unlinkSync(tmp);
      }
    } catch {
      /* ignore */
    }
    throw error;
  }

  // post-commit best-effort rotate — never throws into caller rollback for install success
  try {
    rotateOwnPublishBackups(jsaddons);
  } catch (error) {
    if (opts.collectRotateWarning) {
      opts.collectRotateWarning(
        error instanceof Error ? error.message : String(error),
      );
    }
  }
  return { backedUp, previousBytes, previousExisted };
}

export function restorePublishBytes(jsaddons, publishPath, previousBytes, previousExisted) {
  assertInside(jsaddons, publishPath, "publish.xml");
  const st = lstatIfPresent(publishPath);
  if (st && st.isSymbolicLink()) {
    throw new Error(`cannot restore publish over symlink: ${publishPath}`);
  }
  if (!previousExisted) {
    if (st) {
      if (!st.isFile()) throw new Error(`cannot remove non-file publish: ${publishPath}`);
      fs.unlinkSync(publishPath);
    }
    return;
  }
  fs.writeFileSync(publishPath, previousBytes, "utf8");
}
