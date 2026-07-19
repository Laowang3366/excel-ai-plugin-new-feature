/**
 * Office Excel TaskPane manifest template render + validate.
 * Pure helpers; CLI lives in render-office-manifest.mjs / check-office-manifest.mjs.
 */

export const DEFAULT_APP_ID = "8f3c2a91-6b4e-4d1a-9c77-a1b2c3d4e5f6";
export const DEFAULT_VERSION = "0.1.0.0";
export const DEFAULT_DEV_BASE_URL = "https://localhost:3000";
export const SUPPORT_URL = "https://plugin.shelelove.top";

const UUID_RE =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const VERSION_RE = /^\d+\.\d+\.\d+\.\d+$/;
const PLACEHOLDER_RE = /__\w+__/g;
const UNSAFE_URL_CHAR_RE = /[<>"'&]/;
/** Raw & not starting a known entity → invalid XML attribute. */
const RAW_AMP_RE = /&(?!amp;|lt;|gt;|quot;|apos;|#\d+;|#x[0-9a-fA-F]+;)/i;

export function escapeXmlAttr(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

export function unescapeXmlAttr(value) {
  return String(value)
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function assertSafeAttrRaw(raw, label) {
  if (RAW_AMP_RE.test(raw)) {
    throw new Error(`invalid XML attribute entity in ${label}: ${raw}`);
  }
  return unescapeXmlAttr(raw);
}

/** Match first `name="..."` attribute value with entity checks. */
export function extractAttr(xml, name) {
  const re = new RegExp(`${name}\\s*=\\s*"([^"]*)"`, "i");
  const m = xml.match(re);
  if (!m) return null;
  return assertSafeAttrRaw(m[1], name);
}

function extractTaggedDefault(xml, tagName) {
  const re = new RegExp(`<${tagName}\\s+DefaultValue="([^"]*)"`, "i");
  const m = xml.match(re);
  if (!m) return "";
  return assertSafeAttrRaw(m[1], tagName);
}

function extractResidDefault(xml, tag, id) {
  const re = new RegExp(
    `<${tag}\\s+id="${id}"\\s+DefaultValue="([^"]*)"`,
    "i",
  );
  const m = xml.match(re);
  if (!m) return "";
  return assertSafeAttrRaw(m[1], id);
}

export function normalizeBaseUrl(input) {
  if (typeof input !== "string" || input.trim() === "") {
    throw new Error("baseUrl is required");
  }
  const original = input.trim();
  if (UNSAFE_URL_CHAR_RE.test(original)) {
    throw new Error(
      `baseUrl contains characters unsafe for XML attributes (& < > " '): ${input}`,
    );
  }
  let raw = original.replace(/\/+$/, "");
  let url;
  try {
    url = new URL(raw);
  } catch {
    throw new Error(`Invalid baseUrl: ${input}`);
  }
  if (url.protocol !== "https:") {
    throw new Error(`baseUrl must use HTTPS: ${input}`);
  }
  if (url.username || url.password) {
    throw new Error("baseUrl must not include credentials");
  }
  if (url.search || url.hash) {
    throw new Error("baseUrl must not include query or hash");
  }
  let pathDecoded;
  try {
    pathDecoded = decodeURIComponent(url.pathname);
  } catch {
    throw new Error(`baseUrl path is not valid percent-encoding: ${input}`);
  }
  if (UNSAFE_URL_CHAR_RE.test(pathDecoded) || UNSAFE_URL_CHAR_RE.test(url.pathname)) {
    throw new Error(
      `baseUrl path contains characters unsafe for XML attributes: ${input}`,
    );
  }
  const pathPart = url.pathname === "/" ? "" : url.pathname.replace(/\/+$/, "");
  return `${url.origin}${pathPart}`;
}

export function baseUrlOrigin(baseUrl) {
  return new URL(normalizeBaseUrl(baseUrl)).origin;
}

export function joinBaseUrl(baseUrl, relPath) {
  const base = normalizeBaseUrl(baseUrl);
  const rel = String(relPath || "").replace(/^\/+/, "");
  if (UNSAFE_URL_CHAR_RE.test(rel)) {
    throw new Error(`relative path contains unsafe characters: ${relPath}`);
  }
  return `${base}/${rel}`;
}

export function isLocalhostHost(hostname) {
  const h = String(hostname || "").toLowerCase();
  return h === "localhost" || h === "127.0.0.1" || h === "[::1]" || h === "::1";
}

function urlUnderBase(candidate, baseUrl) {
  try {
    const base = new URL(baseUrl.endsWith("/") ? baseUrl : `${baseUrl}/`);
    const u = new URL(candidate);
    if (u.origin !== base.origin) return false;
    const basePath = base.pathname.endsWith("/") ? base.pathname : `${base.pathname}/`;
    const path = u.pathname.startsWith("/") ? u.pathname : `/${u.pathname}`;
    if (basePath === "/") return true;
    return path === basePath.slice(0, -1) || path.startsWith(basePath);
  } catch {
    return false;
  }
}

/**
 * @param {{ mode: "dev"|"prod", baseUrl?: string, version?: string, appId?: string, template: string }} opts
 */
export function renderOfficeManifest(opts) {
  const mode = opts.mode;
  if (mode !== "dev" && mode !== "prod") {
    throw new Error(`mode must be dev|prod, got: ${mode}`);
  }
  const baseUrl = normalizeBaseUrl(opts.baseUrl ?? DEFAULT_DEV_BASE_URL);
  const version = opts.version ?? DEFAULT_VERSION;
  const appId = opts.appId ?? DEFAULT_APP_ID;
  if (!VERSION_RE.test(version)) {
    throw new Error(`version must be four-part numeric (x.y.z.w), got: ${version}`);
  }
  if (!UUID_RE.test(appId)) {
    throw new Error(`appId must be UUID, got: ${appId}`);
  }

  const origin = baseUrlOrigin(baseUrl);
  if (mode === "prod" && isLocalhostHost(new URL(origin).hostname)) {
    throw new Error("prod baseUrl must not use localhost / 127.0.0.1 / ::1");
  }

  const asset = (p) => joinBaseUrl(baseUrl, p);
  const replacements = {
    __APP_ID__: escapeXmlAttr(appId),
    __VERSION__: escapeXmlAttr(version),
    __BASE_URL__: escapeXmlAttr(baseUrl),
    __APP_DOMAIN__: escapeXmlAttr(origin),
    __SUPPORT_URL__: escapeXmlAttr(SUPPORT_URL),
    __ICON_16__: escapeXmlAttr(asset("assets/icon-16.png")),
    __ICON_32__: escapeXmlAttr(asset("assets/icon-32.png")),
    __ICON_64__: escapeXmlAttr(asset("assets/icon-64.png")),
    __ICON_80__: escapeXmlAttr(asset("assets/icon-80.png")),
    __SOURCE_LOCATION__: escapeXmlAttr(asset("index.html")),
    __COMMANDS_URL__: escapeXmlAttr(asset("index.html")),
    __TASKPANE_URL__: escapeXmlAttr(asset("index.html")),
    __DISPLAY_NAME__: escapeXmlAttr("文格 Excel AI 验证加载项"),
    __DESCRIPTION__: escapeXmlAttr(
      "Excel AI chat + tools development validation add-in (Office.js task pane). Host verification on real Excel/WPS is not claimed.",
    ),
    __GETSTARTED_DESCRIPTION__: escapeXmlAttr(
      "Open the task pane to validate Excel AI chat, read-only queries, and approval-gated workbook tools.",
    ),
  };

  let out = opts.template;
  for (const [key, value] of Object.entries(replacements)) {
    out = out.split(key).join(value);
  }
  const leftover = out.match(PLACEHOLDER_RE);
  if (leftover?.length) {
    throw new Error(`Unresolved template placeholders: ${[...new Set(leftover)].join(", ")}`);
  }
  return out;
}

/**
 * @param {string} xml
 * @param {{ mode?: "dev"|"prod" }} [opts]
 */
export function validateOfficeManifest(xml, opts = {}) {
  const errors = [];
  if (typeof xml !== "string" || xml.trim() === "") {
    return { ok: false, errors: ["manifest is empty"] };
  }
  if (PLACEHOLDER_RE.test(xml)) errors.push("unresolved template placeholders remain");
  PLACEHOLDER_RE.lastIndex = 0;
  if (!xml.includes('xsi:type="TaskPaneApp"')) errors.push("missing xsi:type=TaskPaneApp");

  const id = xml.match(/<Id>([^<]+)<\/Id>/)?.[1] ?? "";
  if (!UUID_RE.test(id)) errors.push(`invalid Id UUID: ${id}`);
  const version = xml.match(/<Version>([^<]+)<\/Version>/)?.[1] ?? "";
  if (!VERSION_RE.test(version)) errors.push(`invalid Version: ${version}`);
  if (!xml.includes("<Permissions>ReadWriteDocument</Permissions>")) {
    errors.push("missing ReadWriteDocument permissions");
  }

  let source = "";
  let commands = "";
  let taskpane = "";
  let icon16 = "";
  let icon32 = "";
  let icon64 = "";
  let icon80 = "";
  let support = "";
  try {
    // Fail closed on any DefaultValue with bare &.
    for (const m of xml.matchAll(/DefaultValue="([^"]*)"/g)) {
      assertSafeAttrRaw(m[1], "DefaultValue");
    }
    source = extractTaggedDefault(xml, "SourceLocation");
    commands = extractResidDefault(xml, "bt:Url", "Commands.Url");
    taskpane = extractResidDefault(xml, "bt:Url", "Taskpane.Url");
    icon16 = extractResidDefault(xml, "bt:Image", "Icon.16");
    icon32 = extractResidDefault(xml, "bt:Image", "Icon.32");
    icon80 = extractResidDefault(xml, "bt:Image", "Icon.80");
    icon64 = extractTaggedDefault(xml, "HighResolutionIconUrl");
    support = extractTaggedDefault(xml, "SupportUrl");
    // IconUrl should match 32 when present.
    const iconUrl = extractTaggedDefault(xml, "IconUrl");
    if (!icon16) icon16 = iconUrl;
    if (!icon32) icon32 = iconUrl;
  } catch (err) {
    return { ok: false, errors: [String(err?.message || err)] };
  }

  if (!source) errors.push("missing SourceLocation");
  if (!commands) errors.push("missing Commands.Url");
  if (!taskpane) errors.push("missing Taskpane.Url");
  if (!icon16?.includes("icon-16.png")) errors.push("missing icon-16");
  if (!icon32?.includes("icon-32.png")) errors.push("missing icon-32");
  if (!icon64?.includes("icon-64.png")) {
    errors.push("missing icon-64 HighResolutionIconUrl");
  }
  if (!icon80?.includes("icon-80.png")) errors.push("missing icon-80");

  const addinUrls = [source, commands, taskpane, icon16, icon32, icon64, icon80].filter(
    Boolean,
  );
  for (const u of addinUrls) {
    try {
      if (new URL(u).protocol !== "https:") errors.push(`non-HTTPS URL: ${u}`);
    } catch {
      errors.push(`invalid URL: ${u}`);
    }
  }

  let baseUrl = "";
  try {
    const su = new URL(source);
    const path = su.pathname.endsWith(".html")
      ? su.pathname.replace(/\/[^/]*$/, "")
      : su.pathname.replace(/\/+$/, "");
    baseUrl = `${su.origin}${path === "/" ? "" : path}`;
  } catch {
    /* reported */
  }

  if (baseUrl) {
    for (const u of addinUrls) {
      if (!urlUnderBase(u, baseUrl)) {
        errors.push(`add-in URL not under base ${baseUrl}: ${u}`);
      }
    }
    const domains = [...xml.matchAll(/<AppDomain>([^<]+)<\/AppDomain>/g)].map((m) => m[1]);
    try {
      const origin = new URL(baseUrl).origin;
      if (!domains.includes(origin)) {
        errors.push(`AppDomains missing SourceLocation origin ${origin}`);
      }
    } catch {
      /* ignore */
    }
  }

  if (support && !support.startsWith("https://")) {
    errors.push(`SupportUrl not HTTPS: ${support}`);
  }
  if (/first batch|首批/i.test(xml)) {
    errors.push("stale first-batch wording remains in manifest");
  }
  if (opts.mode === "prod") {
    for (const u of addinUrls) {
      try {
        if (isLocalhostHost(new URL(u).hostname)) {
          errors.push(`prod forbids localhost host: ${u}`);
        }
      } catch {
        /* ignore */
      }
    }
  }
  return { ok: errors.length === 0, errors };
}
