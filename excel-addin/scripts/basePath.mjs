/**
 * Normalize Vite production/public base path.
 * Accepts "/", "/excel-addin/", "excel-addin"; rejects protocol/query/hash/traversal.
 */
export function normalizeBasePath(input) {
  if (input == null) return "/";
  if (typeof input !== "string") {
    throw new Error("VITE_BASE must be a string path");
  }
  let raw = input.trim();
  if (raw === "") return "/";
  if (/^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(raw) || raw.includes("?") || raw.includes("#")) {
    throw new Error(`Invalid VITE_BASE (absolute URL/query/hash not allowed): ${input}`);
  }
  if (raw.includes("\\")) {
    throw new Error(`Invalid VITE_BASE (backslashes not allowed): ${input}`);
  }
  while (raw.includes("//")) raw = raw.replace(/\/\//g, "/");
  if (!raw.startsWith("/")) raw = `/${raw}`;
  if (!raw.endsWith("/")) raw = `${raw}/`;
  // Reject empty or traversal segments: "/../", "/./", "//" already collapsed.
  const segments = raw.split("/").filter((s) => s.length > 0);
  for (const seg of segments) {
    if (seg === "." || seg === "..") {
      throw new Error(`Invalid VITE_BASE (path traversal not allowed): ${input}`);
    }
    if (!/^[A-Za-z0-9._-]+$/.test(seg)) {
      throw new Error(`Invalid VITE_BASE path segments: ${input}`);
    }
  }
  return raw;
}

export function resolveViteBase(env = process.env) {
  return normalizeBasePath(env.VITE_BASE ?? "/");
}
