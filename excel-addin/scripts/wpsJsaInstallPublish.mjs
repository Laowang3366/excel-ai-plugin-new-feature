/**
 * Safe publish.xml merge for a single WenggeExcelAiAddin jsplugin entry.
 * Fail closed on DOCTYPE/ENTITY, multi-root, nested plugins, duplicates.
 */
import fs from "node:fs";
import path from "node:path";
import {
  WPS_ADDON_NAME,
  WPS_PUBLISH_URL,
} from "./wpsJsaPackage.mjs";
import {
  assertInside,
  assertRealFile,
  ownPublishBackupPath,
  rotateOwnPublishBackups,
  safeRenameInside,
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

function normalizeNewlines(text) {
  return String(text).replace(/\r\n/g, "\n");
}

/**
 * @returns {{ plugins: { raw: string, attrs: Record<string,string> }[], warnings: string[] }}
 */
export function parseJspluginsDocument(xml) {
  const warnings = [];
  const text = normalizeNewlines(xml).trim();
  if (!text) throw new Error("publish.xml is empty");
  if (/<!DOCTYPE|<!ENTITY/i.test(text)) {
    throw new Error("publish.xml must not contain DOCTYPE/ENTITY");
  }
  if ((text.match(/<jsplugins\b/gi) || []).length !== 1) {
    throw new Error("publish.xml must have exactly one jsplugins root");
  }
  if ((text.match(/<\/jsplugins>/gi) || []).length !== 1) {
    throw new Error("publish.xml missing closing jsplugins");
  }
  // Reject nested roots / trailing junk
  const rootMatch = text.match(
    /^(?:\s*<\?xml\b[^?]*\?>\s*)?<jsplugins\b[^>]*>([\s\S]*)<\/jsplugins>\s*$/i,
  );
  if (!rootMatch) {
    throw new Error("publish.xml root structure is not a single jsplugins document");
  }
  if (/\b[^>]*\b/i.test(rootMatch[0].match(/<jsplugins\b[^>]*>/i)?.[0] ?? "") === false) {
    // no-op; attributes on root are ignored if present but we reject non-empty attrs for safety
  }
  const rootOpen = text.match(/<jsplugins\b([^>]*)>/i)?.[1] ?? "";
  if (rootOpen.trim() !== "") {
    throw new Error("publish.xml jsplugins root must not have attributes");
  }

  let inner = rootMatch[1];
  // Reject open-close plugin pairs or nested tags other than self-closing jsplugin
  if (/<jsplugin\b[^>]*>[\s\S]*?<\/jsplugin>/i.test(inner)) {
    throw new Error("publish.xml jsplugin must be self-closing");
  }
  if (/<(?!jsplugin\b)[a-zA-Z!]/.test(inner.replace(/<!--[\s\S]*?-->/g, ""))) {
    // comments not allowed either for simplicity
    throw new Error("publish.xml contains non-jsplugin markup");
  }
  if (inner.includes("<!--")) {
    throw new Error("publish.xml comments are not supported");
  }

  const plugins = [];
  const re = /^\s*(<jsplugin\b[\s\S]*?\/>)\s*/;
  while (inner.trim() !== "") {
    const m = inner.match(re);
    if (!m) {
      throw new Error("publish.xml has malformed jsplugin entry or unsafe content");
    }
    const raw = m[1];
    const attrs = parsePluginAttributes(raw);
    plugins.push({ raw, attrs });
    inner = inner.slice(m[0].length);
  }

  const names = plugins.map((p) => p.attrs.name);
  const own = names.filter((n) => n === WPS_ADDON_NAME);
  if (own.length > 1) {
    throw new Error("publish.xml contains duplicate WenggeExcelAiAddin entries");
  }
  // legacy conflict warning only
  for (const n of names) {
    if (n && n !== WPS_ADDON_NAME && /excelaiwps|excel-ai-wps|wenggeexcel/i.test(n)) {
      warnings.push(`possible legacy/conflicting plugin entry preserved: ${n}`);
    }
  }
  return { plugins, warnings };
}

function parsePluginAttributes(rawTag) {
  const inner = rawTag.replace(/^<jsplugin\b/i, "").replace(/\/>$/, "");
  // Only name="value" pairs and whitespace
  if (/[^a-zA-Z0-9_="\s.]/.test(inner.replace(/="[^"]*"/g, ""))) {
    // allow more in values; check structure via regex extract
  }
  const attrs = {};
  const re = /([A-Za-z_][\w-]*)\s*=\s*"([^"]*)"/g;
  let m;
  let consumed = "";
  const matches = [];
  while ((m = re.exec(inner))) {
    matches.push(m);
    if (Object.prototype.hasOwnProperty.call(attrs, m[1])) {
      throw new Error(`duplicate attribute ${m[1]} on jsplugin`);
    }
    attrs[m[1]] = m[2];
  }
  // Ensure the tag body is only whitespace + attr pairs
  let rest = inner;
  for (const match of matches) {
    const idx = rest.indexOf(match[0]);
    if (idx < 0) throw new Error("jsplugin attribute parse failed");
    const before = rest.slice(0, idx);
    if (before.trim() !== "") {
      throw new Error("jsplugin has unsafe non-attribute content");
    }
    rest = rest.slice(idx + match[0].length);
  }
  if (rest.trim() !== "") {
    throw new Error("jsplugin has trailing unsafe content");
  }
  if (!attrs.name) throw new Error("jsplugin missing name attribute");
  return attrs;
}

export function renderJspluginsDocument(plugins) {
  const body = plugins.map((p) => p.raw).join("\n");
  return `${XML_DECL}\n<jsplugins>\n${body}\n</jsplugins>\n`;
}

/**
 * Upsert own plugin; preserve foreign raw entries order.
 */
export function upsertOwnPlugin(xml) {
  const { plugins, warnings } = parseJspluginsDocument(xml || emptyPublish());
  const ownRaw = renderOwnJsplugin();
  const own = { raw: ownRaw, attrs: parsePluginAttributes(ownRaw.trim()) };
  const next = [];
  let replaced = false;
  for (const p of plugins) {
    if (p.attrs.name === WPS_ADDON_NAME) {
      next.push(own);
      replaced = true;
    } else {
      next.push(p);
    }
  }
  if (!replaced) next.push(own);
  return { xml: renderJspluginsDocument(next), warnings, plugins: next };
}

export function removeOwnPlugin(xml) {
  if (!xml || String(xml).trim() === "") {
    return { xml: emptyPublish(), warnings: [], removed: false };
  }
  const { plugins, warnings } = parseJspluginsDocument(xml);
  const next = plugins.filter((p) => p.attrs.name !== WPS_ADDON_NAME);
  const removed = next.length !== plugins.length;
  return { xml: renderJspluginsDocument(next), warnings, removed };
}

export function emptyPublish() {
  return `${XML_DECL}\n<jsplugins>\n</jsplugins>\n`;
}

export function ownPluginMatchesContract(attrs) {
  return (
    attrs?.name === WPS_ADDON_NAME &&
    attrs?.type === "et" &&
    attrs?.url === WPS_PUBLISH_URL &&
    attrs?.debug === "" &&
    attrs?.enable === "enable_dev"
  );
}

/**
 * Atomic write with own-prefix backup of previous publish.xml.
 */
export function writePublishXmlAtomic(jsaddons, publishPath, newXml) {
  assertInside(jsaddons, publishPath, "publish.xml");
  const tmp = path.join(jsaddons, `publish.xml.wengge-excel-ai.tmp.${Date.now()}`);
  assertInside(jsaddons, tmp, "publish tmp");

  let backedUp = null;
  if (fs.existsSync(publishPath)) {
    assertRealFile(publishPath, "publish.xml");
    backedUp = ownPublishBackupPath(jsaddons);
    assertInside(jsaddons, backedUp, "publish backup");
    fs.copyFileSync(publishPath, backedUp);
    rotateOwnPublishBackups(jsaddons);
  }

  try {
    fs.writeFileSync(tmp, newXml, "utf8");
    const fd = fs.openSync(tmp, "r+");
    try {
      fs.fsyncSync(fd);
    } finally {
      fs.closeSync(fd);
    }
    // atomic replace
    fs.renameSync(tmp, publishPath);
  } catch (error) {
    try {
      if (fs.existsSync(tmp)) fs.unlinkSync(tmp);
    } catch {
      /* ignore */
    }
    if (backedUp && fs.existsSync(backedUp) && !fs.existsSync(publishPath)) {
      fs.copyFileSync(backedUp, publishPath);
    }
    throw error;
  }
  return { backedUp };
}
