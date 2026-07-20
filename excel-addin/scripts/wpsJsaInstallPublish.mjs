/**
 * Restricted deterministic publish.xml tokenizer/parser for jsaddons.
 * Accepts: optional XML declaration + bare <jsplugins> + self-closing <jsplugin .../> only.
 */
import fs from "node:fs";
import path from "node:path";
import { WPS_ADDON_NAME, WPS_PUBLISH_URL } from "./wpsJsaPackage.mjs";
import {
  assertInside,
  assertRealFile,
  exclusiveTempFile,
  ownPublishBackupPath,
  rotateOwnPublishBackups,
  safeRenameInside,
  TMP_PREFIX,
} from "./wpsJsaInstallPaths.mjs";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8"?>';

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
  if (/<!DOCTYPE|<!ENTITY|<!--|<\[CDATA\[|<\?[^x]/i.test(decl.slice(5))) {
    throw new Error("unsafe content in XML declaration");
  }
  // only allow version/encoding style declaration
  if (!/^<\?xml\s+version\s*=\s*(["'])1\.0\1(?:\s+encoding\s*=\s*(["'])[^"']+\2)?\s*\?>$/i.test(decl)) {
    // tolerate encoding=utf-8 casing as real host files
    if (!/^<\?xml\b[^?]*\?>$/i.test(decl) || /<!|\[|\]/.test(decl)) {
      throw new Error("invalid XML declaration");
    }
  }
  return end + 2;
}

/**
 * Parse attributes from inside a jsplugin open segment (without <jsplugin and />).
 * @returns {Record<string,string>}
 */
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
    i += 1; // closing quote
    if (Object.prototype.hasOwnProperty.call(attrs, name)) {
      throw new Error(`duplicate attribute: ${name}`);
    }
    attrs[name] = value;
  }
  return attrs;
}

/**
 * @returns {{ plugins: { raw: string, attrs: Record<string,string> }[], warnings: string[] }}
 */
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
  // reject processing instructions other than leading xml decl
  let i = 0;
  i = parseXmlDeclaration(text, i);
  i = skipWs(text, i);
  if (!text.startsWith("<jsplugins", i)) {
    throw new Error("publish.xml must start with jsplugins root");
  }
  const openEnd = text.indexOf(">", i);
  if (openEnd < 0) throw new Error("unterminated jsplugins open tag");
  const openTag = text.slice(i, openEnd + 1);
  if (openTag.includes("/>")) {
    throw new Error("jsplugins root must not be self-closing empty without body form mismatch");
  }
  // <jsplugins> only — no attributes
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
    // find end of this tag — must be self-closing />
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
        // check self-closing
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
    if (j >= text.length && !text.slice(i).includes(">")) {
      throw new Error("unterminated jsplugin tag");
    }
  }

  i = skipWs(text, i);
  if (i !== text.length) {
    throw new Error("publish.xml has trailing content after jsplugins");
  }

  const names = plugins.map((p) => p.attrs.name);
  const ownCount = names.filter((n) => n === WPS_ADDON_NAME).length;
  if (ownCount > 1) {
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
      // preserve foreign raw tag (trim and re-indent)
      const raw = p.raw.trim();
      return `  ${raw}`;
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

export function ownPluginMatchesContract(attrs) {
  return (
    attrs?.name === WPS_ADDON_NAME &&
    attrs?.type === "et" &&
    attrs?.url === WPS_PUBLISH_URL &&
    attrs?.enable === "enable_dev" &&
    (attrs?.debug === "" || attrs?.debug == null)
  );
}

/**
 * Atomic publish write with own-prefix backup of previous content.
 * @returns {{ backedUp: string|null, previousBytes: string|null, previousExisted: boolean }}
 */
export function writePublishXmlAtomic(jsaddons, publishPath, newXml, opts = {}) {
  assertInside(jsaddons, publishPath, "publish.xml");
  if (path.basename(publishPath) !== "publish.xml") {
    throw new Error("publish path basename must be publish.xml");
  }

  let previousBytes = null;
  let previousExisted = false;
  let backedUp = null;

  if (fs.existsSync(publishPath)) {
    assertRealFile(publishPath, "publish.xml");
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
    rotateOwnPublishBackups(jsaddons);
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
    // replace: if publish exists, rename aside then put new (or overwrite via rename on POSIX)
    if (fs.existsSync(publishPath)) {
      // On POSIX rename over file replaces atomically
      fs.renameSync(tmp, publishPath);
    } else {
      fs.renameSync(tmp, publishPath);
    }
    if (typeof opts.failAfterCommit === "function") opts.failAfterCommit();
  } catch (error) {
    try {
      if (fs.existsSync(tmp)) fs.unlinkSync(tmp);
    } catch {
      /* ignore */
    }
    // restore previous if we may have damaged
    if (previousExisted && previousBytes != null && backedUp && fs.existsSync(backedUp)) {
      try {
        if (!fs.existsSync(publishPath) || fs.readFileSync(publishPath, "utf8") !== previousBytes) {
          fs.writeFileSync(publishPath, previousBytes, "utf8");
        }
      } catch {
        /* best effort; caller may also restore */
      }
    }
    throw error;
  }
  return { backedUp, previousBytes, previousExisted };
}

/** Restore publish.xml bytes or delete if it did not exist. */
export function restorePublishBytes(jsaddons, publishPath, previousBytes, previousExisted) {
  assertInside(jsaddons, publishPath, "publish.xml");
  if (!previousExisted) {
    if (fs.existsSync(publishPath)) {
      const st = fs.lstatSync(publishPath);
      if (st.isSymbolicLink()) throw new Error("cannot restore over symlink publish");
      fs.unlinkSync(publishPath);
    }
    return;
  }
  fs.writeFileSync(publishPath, previousBytes, "utf8");
}
