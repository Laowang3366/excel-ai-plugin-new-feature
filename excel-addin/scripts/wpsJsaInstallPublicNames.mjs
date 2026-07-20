/**
 * Public projection of foreign plugin names for CLI JSON/warnings.
 * Does not rewrite publish.xml; only sanitizes values echoed to operators.
 */

export const UNSAFE_PLUGIN_NAME_PLACEHOLDER = "(unsafe-plugin-name)";
/** Hard cap for names included in public install/status/dry-run JSON. */
export const MAX_PUBLIC_PLUGIN_NAME_LEN = 64;

/**
 * Safe display names: length 1..64, no control/newline chars.
 * Allowed: ASCII letters/digits, space, . _ - and CJK Unified Ideographs.
 * @param {unknown} name
 */
export function isSafePublicPluginName(name) {
  if (typeof name !== "string") return false;
  if (name.length < 1 || name.length > MAX_PUBLIC_PLUGIN_NAME_LEN) return false;
  if (/[\u0000-\u001f\u007f]/.test(name)) return false;
  // \w is [A-Za-z0-9_] in JS without unicode flag for ASCII; allow CJK explicitly.
  return /^[A-Za-z0-9._\- \u4e00-\u9fff]+$/.test(name);
}

/**
 * @param {unknown} name
 * @returns {string}
 */
export function projectPublicPluginName(name) {
  return isSafePublicPluginName(name) ? /** @type {string} */ (name) : UNSAFE_PLUGIN_NAME_PLACEHOLDER;
}

/**
 * Project + dedupe for public lists (unsafe collapses to one placeholder).
 * @param {Iterable<unknown>} names
 * @returns {string[]}
 */
export function projectPublicPluginNames(names) {
  /** @type {string[]} */
  const out = [];
  const seen = new Set();
  for (const raw of names) {
    const projected = projectPublicPluginName(raw);
    if (seen.has(projected)) continue;
    seen.add(projected);
    out.push(projected);
  }
  return out;
}

/**
 * Legacy/third-party warning text without echoing unsafe raw values.
 * @param {unknown} name
 */
export function projectLegacyPluginWarning(name) {
  return `legacy or third-party plugin present: ${projectPublicPluginName(name)}`;
}

/**
 * Sanitize a free-form warning string that may embed a raw plugin name after a known prefix.
 * If no recognized prefix, drop the warning when it contains control chars.
 * @param {unknown} warning
 */
export function projectPublicWarning(warning) {
  if (typeof warning !== "string") return "warning";
  if (/[\u0000-\u001f\u007f]/.test(warning)) {
    return "legacy or third-party plugin present: (unsafe-plugin-name)";
  }
  const prefix = "legacy or third-party plugin present: ";
  if (warning.startsWith(prefix)) {
    return projectLegacyPluginWarning(warning.slice(prefix.length));
  }
  if (warning.length > 240) return `${warning.slice(0, 240)}…`;
  return warning;
}

/**
 * @param {Iterable<unknown>} warnings
 * @returns {string[]}
 */
export function projectPublicWarnings(warnings) {
  /** @type {string[]} */
  const out = [];
  const seen = new Set();
  for (const w of warnings) {
    const projected = projectPublicWarning(w);
    if (seen.has(projected)) continue;
    seen.add(projected);
    out.push(projected);
  }
  return out;
}
