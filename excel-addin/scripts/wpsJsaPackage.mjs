import { assertIndexAssetsUnderBase } from "./packageProdCore.mjs";

export const WPS_ADDON_NAME = "WenggeExcelAiAddin";
/** Canonical jsaddons directory (WPS name + trailing underscore; matches host authaddin path). */
export const WPS_ADDON_DIRECTORY = "WenggeExcelAiAddin_";
/** Phase56–58 mistaken kebab-case layout; install may migrate after verified match. */
export const LEGACY_OWN_ADDON_DIRECTORY = "wengge-excel-ai-addin";
export const WPS_ENTRY_SCRIPT = "wps-entry.js";
export const WPS_PUBLISH_URL =
  `file://%AppData%/kingsoft/wps/jsaddons/${WPS_ADDON_DIRECTORY}/index.html`;
export const LEGACY_OWN_PUBLISH_URL =
  `file://%AppData%/kingsoft/wps/jsaddons/${LEGACY_OWN_ADDON_DIRECTORY}/index.html`;

const OFFICE_JS_URL =
  "https://appsforoffice.microsoft.com/lib/1/hosted/office.js";
const SAFE_VERSION_RE = /^\d+\.\d+\.\d+(?:\.\d+)?$/;
const SAFE_SHA_RE = /^[0-9a-f]{7,64}$/i;
const XML_DANGEROUS_RE = /<!DOCTYPE|<!ENTITY/i;
const PLACEHOLDER_RE = /__\w+__/;

function normalizeText(value) {
  return String(value).replace(/\r\n/g, "\n").trim();
}

function tagText(xml, tagName) {
  const match = xml.match(new RegExp(`<${tagName}>([^<]*)</${tagName}>`, "i"));
  return match?.[1]?.trim() ?? "";
}

function attribute(xml, tagName, name) {
  const tag = xml.match(new RegExp(`<${tagName}\\b[^>]*>`, "i"))?.[0] ?? "";
  const match = tag.match(new RegExp(`\\b${name}\\s*=\\s*"([^"]*)"`, "i"));
  return match?.[1] ?? "";
}

function baseXmlErrors(xml, rootName) {
  const errors = [];
  const text = normalizeText(xml);
  if (!text) errors.push("XML is empty");
  if (XML_DANGEROUS_RE.test(text)) errors.push("DOCTYPE/ENTITY is forbidden");
  if (PLACEHOLDER_RE.test(text)) errors.push("unresolved placeholder remains");
  if (!new RegExp(`<${rootName}\\b`, "i").test(text)) {
    errors.push(`missing ${rootName} root`);
  }
  if (!new RegExp(`</${rootName}>\\s*$`, "i").test(text)) {
    errors.push(`missing closing ${rootName} root`);
  }
  return errors;
}

export function renderWpsPublishXml() {
  return `<?xml version="1.0" encoding="UTF-8"?>
<jsplugins>
  <jsplugin
    name="${WPS_ADDON_NAME}"
    type="et"
    url="${WPS_PUBLISH_URL}"
    debug=""
    enable="enable_dev" />
</jsplugins>
`;
}

export function validateWpsManifest(xml) {
  const errors = baseXmlErrors(xml, "JsPlugin");
  const apiVersion = tagText(xml, "ApiVersion");
  if (!/^\d+\.\d+\.\d+$/.test(apiVersion)) {
    errors.push(`invalid ApiVersion: ${apiVersion}`);
  }
  if (tagText(xml, "Name") !== "Wengge Excel AI Add-in") {
    errors.push("unexpected WPS add-in Name");
  }
  if (!tagText(xml, "Description")) errors.push("Description is required");
  return { ok: errors.length === 0, errors };
}

export function validateWpsRibbon(xml) {
  const errors = baseXmlErrors(xml, "customUI");
  if (!xml.includes('xmlns="http://schemas.microsoft.com/office/2006/01/customui"')) {
    errors.push("missing Office customUI namespace");
  }
  const onLoad = attribute(xml, "customUI", "onLoad");
  if (onLoad !== "WenggeExcelAiOnLoad") {
    errors.push("customUI onLoad must be WenggeExcelAiOnLoad");
  }
  const tab = {
    id: attribute(xml, "tab", "id"),
    visible: attribute(xml, "tab", "getVisible"),
  };
  if (tab.id !== "wenggeExcelAiTab") errors.push("unexpected ribbon tab id");
  if (tab.visible !== "WenggeExcelAiTabVisible") {
    errors.push("ribbon tab getVisible callback mismatch");
  }
  if (attribute(xml, "group", "id") !== "wenggeExcelAiGroup") {
    errors.push("missing WPS ribbon group");
  }

  const buttonTags = xml.match(/<button\b[^>]*>/gi) ?? [];
  const expectedButtons = [
    {
      id: "wenggeExcelAiOpenChatButton",
      onAction: "WenggeExcelAiOpenChat",
    },
    {
      id: "wenggeExcelAiOpenProvidersButton",
      onAction: "WenggeExcelAiOpenProviders",
    },
    {
      id: "wenggeExcelAiOpenHostButton",
      onAction: "WenggeExcelAiOpenHost",
    },
  ];
  for (const expected of expectedButtons) {
    const tag = buttonTags.find((item) =>
      new RegExp(`\\bid\\s*=\\s*"${expected.id}"`, "i").test(item),
    );
    if (!tag) {
      errors.push(`missing WPS ribbon button ${expected.id}`);
      continue;
    }
    const action = tag.match(/\bonAction\s*=\s*"([^"]*)"/i)?.[1] ?? "";
    if (action !== expected.onAction) {
      errors.push(`ribbon button ${expected.id} onAction mismatch`);
    }
    const getImage = tag.match(/\bgetImage\s*=\s*"([^"]*)"/i)?.[1] ?? "";
    if (getImage !== "WenggeExcelAiGetImage") {
      errors.push(`ribbon button ${expected.id} must use getImage=WenggeExcelAiGetImage`);
    }
    if (/\bimage\s*=/.test(tag)) {
      errors.push(`ribbon button ${expected.id} must not use direct image attribute`);
    }
  }
  const remoteUrls = xml.match(/https?:\/\/[^"'\s>]+/gi) ?? [];
  for (const url of remoteUrls) {
    if (!/^https?:\/\/schemas\.microsoft\.com\//i.test(url)) {
      errors.push("remote URLs are forbidden in ribbon.xml");
      break;
    }
  }
  return { ok: errors.length === 0, errors };
}

export function validateWpsEntryScript(source) {
  const errors = [];
  const text = String(source);
  for (const callback of [
    "WenggeExcelAiOnLoad",
    "WenggeExcelAiTabVisible",
    "WenggeExcelAiGetImage",
    "WenggeExcelAiOpenChat",
    "WenggeExcelAiOpenProviders",
    "WenggeExcelAiOpenHost",
  ]) {
    if (!new RegExp(`window\\.${callback}\\s*=\\s*function\\s*\\(`).test(text)) {
      errors.push(`missing global callback ${callback}`);
    }
  }
  if (!/CreateTaskPane/.test(text)) {
    errors.push("WPS entry must call CreateTaskPane");
  }
  if (!/GetTaskPane/.test(text)) {
    errors.push("WPS entry must support GetTaskPane reuse");
  }
  if (!/PluginStorage/.test(text)) {
    errors.push("WPS entry must use PluginStorage for pane id reuse");
  }
  if (!text.includes("assets/icon-32.png") || !text.includes("assets/icon-16.png")) {
    errors.push("WPS entry getImage must map to package-relative icon-32/icon-16 assets");
  }
  if (/\beval\s*\(|\bnew\s+Function\s*\(/.test(text)) {
    errors.push("dynamic code execution is forbidden in WPS entry");
  }
  if (
    /\brequire\s*\(|\bimport\s*\(|from\s+["']electron["']|node:child_process|child_process|\bprocess\.\w+/.test(
      text,
    )
  ) {
    errors.push("Node/Electron/process APIs are forbidden in WPS entry");
  }
  if (/https?:\/\/(?!appsforoffice\.microsoft\.com)/i.test(text) && /https?:\/\/[a-z0-9.-]+\//i.test(text)) {
    // Allow comments? Prefer fail if hard-coded remote task pane hosts appear as string literals.
    const remoteLiterals = text.match(/["']https?:\/\/[^"']+["']/gi) ?? [];
    if (remoteLiterals.some((item) => !/appsforoffice\.microsoft\.com/i.test(item))) {
      errors.push("hard-coded remote URLs are forbidden in WPS entry");
    }
  }
  return { ok: errors.length === 0, errors };
}

export function validateWpsPublishXml(xml) {
  const errors = baseXmlErrors(xml, "jsplugins");
  if (attribute(xml, "jsplugin", "name") !== WPS_ADDON_NAME) {
    errors.push("publish.xml add-in name mismatch");
  }
  if (attribute(xml, "jsplugin", "type") !== "et") {
    errors.push("publish.xml must target WPS spreadsheets (et)");
  }
  const url = attribute(xml, "jsplugin", "url");
  if (url !== WPS_PUBLISH_URL) errors.push("publish.xml URL mismatch");
  if (url.includes("..") || /^https?:/i.test(url)) {
    errors.push("publish.xml URL must remain inside the local jsaddons directory");
  }
  if (attribute(xml, "jsplugin", "enable") !== "enable_dev") {
    errors.push("publish.xml enable mode mismatch");
  }
  return { ok: errors.length === 0, errors };
}

export function prepareWpsIndexHtml(html) {
  const officeScript = new RegExp(
    `<script\\s+src=["']${OFFICE_JS_URL.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}["']\\s*><\\/script>\\s*`,
    "gi",
  );
  const matches = String(html).match(officeScript) ?? [];
  if (matches.length !== 1) {
    throw new Error(`expected exactly one Office.js script, found ${matches.length}`);
  }
  let out = String(html).replace(officeScript, "");
  if (out.includes(WPS_ENTRY_SCRIPT)) throw new Error("WPS entry script already injected");
  if (!/<\/head>/i.test(out)) throw new Error("index.html is missing </head>");
  out = out.replace(
    /<\/head>/i,
    `  <script src="./${WPS_ENTRY_SCRIPT}"></script>\n  </head>`,
  );
  return out;
}

export function validateWpsIndexHtml(html) {
  const errors = [];
  const text = String(html);
  if (text.includes(OFFICE_JS_URL)) errors.push("WPS entry must not load Office.js CDN");
  if (!text.includes(`src="./${WPS_ENTRY_SCRIPT}"`)) {
    errors.push("WPS entry script is not loaded relatively");
  }
  for (const match of text.matchAll(
    /<(?:script|link)\b[^>]*\b(?:src|href)\s*=\s*(["'])(.*?)\1/gi,
  )) {
    const ref = match[2] ?? "";
    if (ref.startsWith("/") && !ref.startsWith("//")) {
      errors.push(`root-absolute asset is invalid for file:// WPS package: ${ref}`);
    }
  }
  let assets = [];
  try {
    assets = assertIndexAssetsUnderBase(text, "/");
  } catch (error) {
    errors.push(error instanceof Error ? error.message : String(error));
  }
  return { ok: errors.length === 0, errors, assets };
}

export function normalizeWpsGitSha(value) {
  const sha = String(value || "").trim();
  return SAFE_SHA_RE.test(sha) ? sha : "unknown";
}

export function makeWpsArtifactName(version, gitSha) {
  const normalizedVersion = String(version || "").trim();
  if (!SAFE_VERSION_RE.test(normalizedVersion)) {
    throw new Error(`invalid package version: ${version}`);
  }
  const shortSha = normalizeWpsGitSha(gitSha).slice(0, 7) || "unknown";
  return `excel-addin-wps-jsa-${normalizedVersion}-${shortSha}`;
}

export function validateWpsSourceBundle(bundle) {
  const checks = [
    validateWpsManifest(bundle.manifestXml),
    validateWpsRibbon(bundle.ribbonXml),
    validateWpsEntryScript(bundle.entryScript),
    validateWpsPublishXml(bundle.publishXml),
  ];
  const errors = checks.flatMap((check) => check.errors);
  const expectedPublish = normalizeText(renderWpsPublishXml());
  if (normalizeText(bundle.publishXml) !== expectedPublish) {
    errors.push("checked-in publish.xml drifts from deterministic render");
  }
  return { ok: errors.length === 0, errors };
}
