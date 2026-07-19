/**
 * Pure helpers for Excel add-in production packaging.
 * No vite spawn / network — CLI in package-prod.mjs orchestrates build.
 */
import { createHash } from "node:crypto";
import { normalizeBasePath } from "./basePath.mjs";
import { isLocalhostHost, normalizeBaseUrl } from "./officeManifest.mjs";

const SEMVER3_RE = /^\d+\.\d+\.\d+$/;
const SEMVER4_RE = /^\d+\.\d+\.\d+\.\d+$/;
const ARTIFACT_SAFE_RE = /^[A-Za-z0-9._-]+$/;

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

/**
 * Local script/link href/src under same origin path must start with viteBase (or be absolute CDN).
 * Returns list of local asset paths found.
 */
export function collectLocalAssetRefs(html) {
  const refs = [];
  const re = /\b(?:src|href)=["']([^"']+)["']/gi;
  let m;
  while ((m = re.exec(html)) !== null) {
    const ref = m[1];
    if (!ref || ref.startsWith("data:") || ref.startsWith("blob:")) continue;
    if (/^https?:\/\//i.test(ref)) {
      // external CDN — not local package asset
      continue;
    }
    refs.push(ref);
  }
  return refs;
}

export function assertIndexAssetsUnderBase(html, viteBase) {
  const base = normalizeBasePath(viteBase);
  const refs = collectLocalAssetRefs(html);
  const errors = [];
  for (const ref of refs) {
    // absolute path on same host
    if (ref.startsWith("/")) {
      if (base === "/") {
        // any absolute path ok under root deploy
        continue;
      }
      if (!(ref === base.slice(0, -1) || ref.startsWith(base))) {
        errors.push(`asset not under VITE_BASE ${base}: ${ref}`);
      }
      continue;
    }
    // relative refs are allowed (resolved under base at serve time)
  }
  if (errors.length) throw new Error(errors.join("; "));
  return refs;
}

export function isSensitiveRelativePath(relPath) {
  const norm = String(relPath).replace(/\\/g, "/");
  if (norm.includes("..")) return true;
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

export function parseCliArgs(argv) {
  const out = {
    baseUrl: null,
    version: null,
    viteBase: null,
    gitSha: null,
    skipBuild: false,
    distDir: null,
    rootDir: null,
  };
  for (let i = 0; i < argv.length; i += 1) {
    const a = argv[i];
    if (a === "--base-url") out.baseUrl = argv[++i];
    else if (a === "--version") out.version = argv[++i];
    else if (a === "--vite-base") out.viteBase = argv[++i];
    else if (a === "--git-sha") out.gitSha = argv[++i];
    else if (a === "--skip-build") out.skipBuild = true;
    else if (a === "--dist-dir") out.distDir = argv[++i];
    else if (a === "--root-dir") out.rootDir = argv[++i];
    else if (a === "-h" || a === "--help") out.help = true;
    else throw new Error(`Unknown arg: ${a}`);
  }
  return out;
}
