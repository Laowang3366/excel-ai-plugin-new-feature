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

export function normalizeBaseUrl(input) {
  if (typeof input !== "string" || input.trim() === "") {
    throw new Error("baseUrl is required");
  }
  let raw = input.trim();
  raw = raw.replace(/\/+$/, "");
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
  const pathPart = url.pathname === "/" ? "" : url.pathname.replace(/\/+$/, "");
  return `${url.origin}${pathPart}`;
}

export function baseUrlOrigin(baseUrl) {
  const normalized = normalizeBaseUrl(baseUrl);
  return new URL(normalized).origin;
}

export function joinBaseUrl(baseUrl, relPath) {
  const base = normalizeBaseUrl(baseUrl);
  const rel = String(relPath || "").replace(/^\/+/, "");
  return `${base}/${rel}`;
}

export function isLocalhostHost(hostname) {
  const h = String(hostname || "").toLowerCase();
  return h === "localhost" || h === "127.0.0.1" || h === "[::1]" || h === "::1";
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
    __APP_ID__: appId,
    __VERSION__: version,
    __BASE_URL__: baseUrl,
    __APP_DOMAIN__: origin,
    __SUPPORT_URL__: SUPPORT_URL,
    __ICON_16__: asset("assets/icon-16.png"),
    __ICON_32__: asset("assets/icon-32.png"),
    __ICON_64__: asset("assets/icon-64.png"),
    __ICON_80__: asset("assets/icon-80.png"),
    __SOURCE_LOCATION__: asset("index.html"),
    __COMMANDS_URL__: asset("index.html"),
    __TASKPANE_URL__: asset("index.html"),
    __DISPLAY_NAME__: "文格 Excel AI 验证加载项",
    __DESCRIPTION__:
      "Excel AI chat + tools development validation add-in (Office.js task pane). Host verification on real Excel/WPS is not claimed.",
    __GETSTARTED_DESCRIPTION__:
      "Open the task pane to validate Excel AI chat, read-only queries, and approval-gated workbook tools.",
  };

  let out = opts.template;
  for (const [key, value] of Object.entries(replacements)) {
    out = out.split(key).join(value);
  }
  const leftover = out.match(PLACEHOLDER_RE);
  if (leftover && leftover.length) {
    throw new Error(`Unresolved template placeholders: ${[...new Set(leftover)].join(", ")}`);
  }
  return out;
}

/**
 * Validate a rendered Office manifest XML string.
 * @param {string} xml
 * @param {{ mode?: "dev"|"prod" }} [opts]
 */
export function validateOfficeManifest(xml, opts = {}) {
  const errors = [];
  if (typeof xml !== "string" || xml.trim() === "") {
    return { ok: false, errors: ["manifest is empty"] };
  }
  if (PLACEHOLDER_RE.test(xml)) {
    errors.push("unresolved template placeholders remain");
  }
  // reset lastIndex after global test
  PLACEHOLDER_RE.lastIndex = 0;
  if (!xml.includes('xsi:type="TaskPaneApp"')) {
    errors.push("missing xsi:type=TaskPaneApp");
  }
  const id = xml.match(/<Id>([^<]+)<\/Id>/)?.[1] ?? "";
  if (!UUID_RE.test(id)) errors.push(`invalid Id UUID: ${id}`);
  const version = xml.match(/<Version>([^<]+)<\/Version>/)?.[1] ?? "";
  if (!VERSION_RE.test(version)) errors.push(`invalid Version: ${version}`);
  if (!xml.includes("<Permissions>ReadWriteDocument</Permissions>")) {
    errors.push("missing ReadWriteDocument permissions");
  }

  const source = xml.match(/<SourceLocation\s+DefaultValue="([^"]+)"/)?.[1] ?? "";
  let sourceOrigin = "";
  try {
    const u = new URL(source);
    if (u.protocol !== "https:") errors.push(`SourceLocation not HTTPS: ${source}`);
    sourceOrigin = u.origin;
  } catch {
    errors.push(`invalid SourceLocation: ${source}`);
  }

  const httpsAttrs = [...xml.matchAll(/DefaultValue="(https?:\/\/[^"]+)"/g)].map((m) => m[1]);
  for (const u of httpsAttrs) {
    if (!u.startsWith("https://")) errors.push(`non-HTTPS URL: ${u}`);
  }

  const domains = [...xml.matchAll(/<AppDomain>([^<]+)<\/AppDomain>/g)].map((m) => m[1]);
  if (sourceOrigin && !domains.includes(sourceOrigin)) {
    errors.push(
      `AppDomains missing SourceLocation origin ${sourceOrigin} (have: ${domains.join(", ")})`,
    );
  }

  for (const needle of [
    "Commands.Url",
    "Taskpane.Url",
    "icon-16.png",
    "icon-32.png",
    "icon-80.png",
  ]) {
    if (!xml.includes(needle)) errors.push(`missing resource reference: ${needle}`);
  }

  if (/first batch|首批/i.test(xml)) {
    errors.push("stale first-batch wording remains in manifest");
  }

  if (opts.mode === "prod") {
    for (const u of httpsAttrs) {
      try {
        const host = new URL(u).hostname;
        if (isLocalhostHost(host)) errors.push(`prod forbids localhost host: ${u}`);
      } catch {
        errors.push(`invalid URL: ${u}`);
      }
    }
  }

  return { ok: errors.length === 0, errors };
}
