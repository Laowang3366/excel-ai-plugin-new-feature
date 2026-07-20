/**
 * Helpers for Excel add-in production packaging.
 * No Vite spawn or network; filesystem checks stay fail-closed.
 */
import { createHash } from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { normalizeBasePath } from "./basePath.mjs";
import { isLocalhostHost, normalizeBaseUrl } from "./officeManifest.mjs";

const SEMVER3_RE = /^\d+\.\d+\.\d+$/;
const SEMVER4_RE = /^\d+\.\d+\.\d+\.\d+$/;
const ARTIFACT_SAFE_RE = /^[A-Za-z0-9._-]+$/;
const SCHEME_RE = /^[A-Za-z][A-Za-z0-9+.-]*:/;
const CONTROL_CHAR_RE = /[\u0000-\u001f\u007f]/;
const ENCODED_SEPARATOR_RE = /%(?:2f|5c)/i;
const PACKAGE_ORIGIN = "https://package.invalid";
const OFFICE_JS_CDN_URL =
  "https://appsforoffice.microsoft.com/lib/1/hosted/office.js";

const SENSITIVE_NAME_RE =
  /(^|\/)(\.env(\..*)?|CLAUDE\.md|.*\.(pem|key|p12|pfx)|node_modules)(\/|$)/i;

/**
 * package.json "0.1.0" → "0.1.0.0"; already four-part left as-is.
 */
export function expandPackageVersion(version) {
  const v = String(version || "").trim();
  if (SEMVER4_RE.test(v)) return v;
  if (SEMVER3_RE.test(v)) return `${v}.0`;
  throw new Error(`package version must be x.y.z or x.y.z.w, got: ${version}`);
}

export function requireFourPartVersion(version) {
  const v = String(version || "").trim();
  if (!SEMVER4_RE.test(v)) {
    throw new Error(`version must be four-part numeric (x.y.z.w), got: ${version}`);
  }
  return v;
}

/** Pathname of baseUrl → Vite base with trailing slash (root → "/"). */
export function deriveViteBaseFromBaseUrl(baseUrl) {
  const normalized = normalizeBaseUrl(baseUrl);
  const { pathname } = new URL(normalized);
  if (!pathname || pathname === "/") return "/";
  return normalizeBasePath(pathname);
}

export function assertViteBaseMatchesBaseUrl(viteBase, baseUrl) {
  const expected = deriveViteBaseFromBaseUrl(baseUrl);
  const actual = normalizeBasePath(viteBase);
  if (actual !== expected) {
    throw new Error(
      `vite_base ${actual} does not match base_url path ${expected} (from ${baseUrl})`,
    );
  }
  return actual;
}

function pathIsInside(root, candidate) {
  const relative = path.relative(root, candidate);
  return (
    relative !== "" &&
    relative !== ".." &&
    !relative.startsWith(`..${path.sep}`) &&
    !path.isAbsolute(relative)
  );
}

function decodeAssetPath(rawPath, ref) {
  if (ENCODED_SEPARATOR_RE.test(rawPath)) {
    throw new Error(`encoded path separators are not allowed in asset URL: ${ref}`);
  }
  let decoded;
  try {
    decoded = decodeURIComponent(rawPath);
  } catch {
    throw new Error(`asset URL has invalid percent encoding: ${ref}`);
  }
  if (decoded.includes("\\") || CONTROL_CHAR_RE.test(decoded)) {
    throw new Error(`asset URL contains backslash/control characters: ${ref}`);
  }
  if (decoded.split("/").includes("..")) {
    throw new Error(`asset URL contains path traversal: ${ref}`);
  }
  return decoded;
}

function collectIndexAssetRefs(html) {
  const refs = [];
  const startRe = /<(script|link)\b/gi;
  const attributeRe =
    /([^\s"'<>/=]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'=<>`]+)))?/g;
  let match;
  while ((match = startRe.exec(html)) !== null) {
    const tagName = match[1].toLowerCase();
    let quote = null;
    let end = -1;
    for (let i = match.index; i < html.length; i += 1) {
      const char = html[i];
      if (quote) {
        if (char === quote) quote = null;
      } else if (char === '"' || char === "'") {
        quote = char;
      } else if (char === ">") {
        end = i + 1;
        break;
      }
    }
    if (end < 0) throw new Error(`unterminated ${tagName} tag in index.html`);
    const attributes = html.slice(match.index + match[0].length, end - 1);
    startRe.lastIndex = end;
    const attrName = tagName === "script" ? "src" : "href";
    let ref = null;
    attributeRe.lastIndex = 0;
    let attribute;
    while ((attribute = attributeRe.exec(attributes)) !== null) {
      if (attribute[1].toLowerCase() !== attrName) continue;
      if (ref !== null) {
        throw new Error(`duplicate ${tagName} ${attrName} is forbidden`);
      }
      if (attribute[2] === undefined && attribute[3] === undefined) {
        throw new Error(`${tagName} ${attrName} must use a quoted URL`);
      }
      ref = attribute[2] ?? attribute[3];
    }
    if (ref !== null) refs.push(ref);
  }
  return refs;
}

function resolveIndexAssetRef(ref, viteBase) {
  if (typeof ref !== "string" || ref === "" || ref !== ref.trim()) {
    throw new Error(`asset URL must be a non-empty trimmed string: ${ref}`);
  }
  if (ref.includes("\\") || CONTROL_CHAR_RE.test(ref)) {
    throw new Error(`asset URL contains backslash/control characters: ${ref}`);
  }
  if (ref.includes("&")) {
    throw new Error(`HTML character references are forbidden in asset URL: ${ref}`);
  }
  if (ref.startsWith("//")) {
    throw new Error(`protocol-relative asset URL is forbidden: ${ref}`);
  }

  if (SCHEME_RE.test(ref)) {
    let external;
    try {
      external = new URL(ref);
    } catch {
      throw new Error(`invalid asset URL: ${ref}`);
    }
    if (external.protocol !== "https:") {
      throw new Error(`unsupported asset URL protocol ${external.protocol}: ${ref}`);
    }
    if (external.href !== OFFICE_JS_CDN_URL) {
      throw new Error(`cross-origin asset URL is not allowlisted: ${ref}`);
    }
    return { kind: "external", source: ref, url: external.href };
  }

  const rawPath = ref.split(/[?#]/, 1)[0];
  decodeAssetPath(rawPath, ref);
  const base = normalizeBasePath(viteBase);
  const documentUrl = new URL(`${base}index.html`, PACKAGE_ORIGIN);
  let resolved;
  try {
    resolved = new URL(ref, documentUrl);
  } catch {
    throw new Error(`invalid local asset URL: ${ref}`);
  }
  if (resolved.origin !== PACKAGE_ORIGIN || resolved.protocol !== "https:") {
    throw new Error(`local asset URL escaped package origin: ${ref}`);
  }

  const decodedPathname = decodeAssetPath(resolved.pathname, ref);
  if (base !== "/" && !decodedPathname.startsWith(base)) {
    throw new Error(`asset not under VITE_BASE ${base}: ${ref}`);
  }
  const relativePath =
    base === "/"
      ? decodedPathname.replace(/^\/+/, "")
      : decodedPathname.slice(base.length);
  if (!relativePath || relativePath.startsWith("/") || relativePath.split("/").includes("..")) {
    throw new Error(`asset URL does not resolve to a package file: ${ref}`);
  }
  return {
    kind: "local",
    source: ref,
    pathname: decodedPathname,
    relativePath,
  };
}

/**
 * Resolve packaging inputs. Fails closed on localhost prod base, mismatches, bad versions.
 * @param {{ baseUrl: string, version?: string|null, viteBase?: string|null, packageJsonVersion: string }} input
 */
export function resolvePackageInputs(input) {
  const baseUrl = normalizeBaseUrl(input.baseUrl);
  if (isLocalhostHost(new URL(baseUrl).hostname)) {
    throw new Error(`production baseUrl must not use localhost: ${input.baseUrl}`);
  }
  const derived = deriveViteBaseFromBaseUrl(baseUrl);
  let viteBase = derived;
  if (input.viteBase != null && String(input.viteBase).trim() !== "") {
    viteBase = assertViteBaseMatchesBaseUrl(input.viteBase, baseUrl);
  }
  let version;
  if (input.version != null && String(input.version).trim() !== "") {
    version = requireFourPartVersion(input.version);
  } else {
    version = expandPackageVersion(input.packageJsonVersion);
  }
  return { baseUrl, viteBase, version, packageJsonVersion: String(input.packageJsonVersion) };
}

export function assertIndexAssetsUnderBase(html, viteBase) {
  const localAssets = [];
  for (const ref of collectIndexAssetRefs(html)) {
    const resolved = resolveIndexAssetRef(ref, viteBase);
    if (resolved.kind === "local") localAssets.push(resolved.relativePath);
  }
  return [...new Set(localAssets)];
}

export function assertLocalAssetFiles(distDir, relativePaths) {
  const rootPath = path.resolve(distDir);
  const rootStat = fs.lstatSync(rootPath);
  if (rootStat.isSymbolicLink() || !rootStat.isDirectory()) {
    throw new Error(`dist root must be a real directory: ${distDir}`);
  }
  const rootReal = fs.realpathSync(rootPath);
  const verified = [];
  for (const relativePath of relativePaths) {
    const normalized = String(relativePath).replace(/\\/g, "/");
    const segments = normalized.split("/");
    if (
      normalized === "" ||
      normalized.startsWith("/") ||
      segments.some((segment) => segment === "" || segment === "." || segment === "..")
    ) {
      throw new Error(`invalid local asset path: ${relativePath}`);
    }

    let current = rootPath;
    for (let i = 0; i < segments.length; i += 1) {
      current = path.join(current, segments[i]);
      let stat;
      try {
        stat = fs.lstatSync(current);
      } catch (error) {
        if (error?.code === "ENOENT") {
          throw new Error(`referenced local asset is missing: ${relativePath}`);
        }
        throw error;
      }
      if (stat.isSymbolicLink()) {
        throw new Error(`referenced local asset uses a symlink: ${relativePath}`);
      }
      if (i < segments.length - 1 && !stat.isDirectory()) {
        throw new Error(`local asset parent is not a directory: ${relativePath}`);
      }
      if (i === segments.length - 1 && !stat.isFile()) {
        throw new Error(`referenced local asset is not a regular file: ${relativePath}`);
      }
    }

    const real = fs.realpathSync(current);
    if (!pathIsInside(rootReal, real)) {
      throw new Error(`referenced local asset escaped dist: ${relativePath}`);
    }
    verified.push(real);
  }
  return verified;
}

export function listFilesRecursiveStrict(dir, base = dir) {
  const root = path.resolve(base);
  const currentDir = path.resolve(dir);
  const rootStat = fs.lstatSync(root);
  if (rootStat.isSymbolicLink() || !rootStat.isDirectory()) {
    throw new Error(`package root must be a real directory: ${base}`);
  }
  const out = [];
  for (const ent of fs.readdirSync(currentDir, { withFileTypes: true })) {
    const abs = path.join(currentDir, ent.name);
    const rel = path.relative(root, abs).replace(/\\/g, "/");
    if (ent.isSymbolicLink()) {
      throw new Error(`symlink is forbidden in package: ${rel}`);
    }
    if (ent.isDirectory()) out.push(...listFilesRecursiveStrict(abs, root));
    else if (ent.isFile()) out.push(rel);
    else throw new Error(`non-regular file is forbidden in package: ${rel}`);
  }
  return out;
}

export function isSensitiveRelativePath(relPath) {
  const norm = String(relPath).replace(/\\/g, "/");
  if (norm.split("/").some((segment) => segment === "." || segment === "..")) {
    return true;
  }
  return SENSITIVE_NAME_RE.test(norm);
}

export function assertNoSensitiveDistPaths(relativePaths) {
  const bad = [];
  for (const p of relativePaths) {
    if (isSensitiveRelativePath(p)) bad.push(p);
  }
  if (bad.length) {
    throw new Error(`sensitive or forbidden paths in dist: ${bad.join(", ")}`);
  }
}

/**
 * @param {Array<{ relativePath: string, content: Buffer|string }>} files
 * @returns {string} SHA256SUMS.txt body
 */
export function buildSha256Sums(files) {
  const rows = files
    .map((f) => {
      const rel = f.relativePath.replace(/\\/g, "/").replace(/^\.\//, "");
      const buf = Buffer.isBuffer(f.content) ? f.content : Buffer.from(f.content);
      const hash = createHash("sha256").update(buf).digest("hex");
      return { rel, hash };
    })
    .filter((r) => r.rel !== "SHA256SUMS.txt")
    .sort((a, b) => (a.rel < b.rel ? -1 : a.rel > b.rel ? 1 : 0));
  return rows.map((r) => `${r.hash}  ${r.rel}`).join("\n") + (rows.length ? "\n" : "");
}

export function buildBuildInfo(meta) {
  return {
    gitSha: meta.gitSha,
    packageVersion: meta.packageVersion,
    manifestVersion: meta.manifestVersion,
    baseUrl: meta.baseUrl,
    viteBase: meta.viteBase,
  };
}

/** Artifact name: excel-addin-<version>-<shortSha> — only safe chars. */
export function makeArtifactName(version, gitSha) {
  const v = requireFourPartVersion(version);
  const short = String(gitSha || "unknown").replace(/[^0-9a-fA-F]/g, "").slice(0, 7) || "unknown";
  const name = `excel-addin-${v}-${short}`;
  if (!ARTIFACT_SAFE_RE.test(name)) {
    throw new Error(`unsafe artifact name: ${name}`);
  }
  return name;
}

export function formatSpawnFailure(result) {
  if (result?.error) {
    const code = result.error.code ? ` ${result.error.code}` : "";
    return `npm run build failed to start${code}`;
  }
  if (result?.signal) {
    return `npm run build terminated by signal ${result.signal}`;
  }
  return `npm run build failed with status ${result?.status ?? "unknown"}`;
}

export function parseCliArgs(argv) {
  const out = {
    baseUrl: null,
    version: null,
    viteBase: null,
    gitSha: null,
  };
  for (let i = 0; i < argv.length; i += 1) {
    const a = argv[i];
    if (a === "--base-url") out.baseUrl = argv[++i];
    else if (a === "--version") out.version = argv[++i];
    else if (a === "--vite-base") out.viteBase = argv[++i];
    else if (a === "--git-sha") out.gitSha = argv[++i];
    else if (a === "-h" || a === "--help") out.help = true;
    else throw new Error(`Unknown arg: ${a}`);
  }
  return out;
}
