/**
 * Runtime desktop-bridge dependency detection for the Excel add-in.
 *
 * Distinguishes **runtime imports/requires** from **documentation / prompt text**
 * that may mention desktop/electron source paths (e.g. synced prompt provenance).
 *
 * Build-time Node CLIs (package-*.mjs) may use child_process; they must not ship
 * inside task-pane bundles and are excluded by package scanners via path filters.
 */

/** Precise runtime import/require patterns only — not bare-word "electron"/"desktop". */
export const RUNTIME_DESKTOP_DEP_PATTERNS = [
  { id: "import-from-electron", re: /\bfrom\s+["']electron["']/ },
  { id: "require-electron", re: /\brequire\(\s*["']electron["']\s*\)/ },
  { id: "dynamic-import-electron", re: /\bimport\(\s*["']electron["']\s*\)/ },
  { id: "import-from-child_process", re: /\bfrom\s+["'](?:node:)?child_process["']/ },
  { id: "require-child_process", re: /\brequire\(\s*["'](?:node:)?child_process["']\s*\)/ },
  { id: "import-from-desktop-path", re: /\bfrom\s+["'][^"']*desktop\// },
  { id: "require-desktop-path", re: /\brequire\(\s*["'][^"']*desktop\// },
  { id: "wengge-office-worker", re: /Wengge\.OfficeWorker/ },
  { id: "system-runtime-interop", re: /System\.Runtime\.InteropServices/ },
  { id: "office-interop", re: /Microsoft\.Office\.Interop/ },
  { id: "edge-js-require", re: /\b(?:require|from)\(?\s*["']edge-js["']/i },
  { id: "node-adodb", re: /\b(?:require|from)\(?\s*["']node-adodb["']/i },
  { id: "dotnet-scope", re: /@dotnet\// },
];

/**
 * @param {string} text
 * @param {string} [label]
 * @returns {{ label: string, id: string }[]}
 */
export function findRuntimeDesktopDepHits(text, label = "snippet") {
  const hits = [];
  for (const { id, re } of RUNTIME_DESKTOP_DEP_PATTERNS) {
    re.lastIndex = 0;
    if (re.test(text)) hits.push({ label, id });
  }
  return hits;
}

/**
 * True when text mentions desktop/electron only as documentation/prompt prose,
 * without runtime import/require forms above.
 */
export function isDocumentationOnlyDesktopMention(text) {
  if (findRuntimeDesktopDepHits(text).length > 0) return false;
  return /desktop\/electron|electron\/agent|COM\s*\/\s*\.NET|child_process/i.test(text);
}

/**
 * Scan package text files for runtime desktop deps.
 * @param {{ relativePath: string, content: string }[]} files
 * @returns {string[]} human-readable offenders
 */
export function collectRuntimeDesktopDepOffenders(files) {
  const offenders = [];
  for (const file of files) {
    const rel = file.relativePath.replace(/\\/g, "/");
    // Never treat hashes/binary as code; callers should only pass text artifacts.
    if (!/\.(js|mjs|cjs|ts|tsx|html|css|json|xml|md|txt)$/i.test(rel)) continue;
    // Build metadata may name tools; still forbid real import forms inside them.
    const hits = findRuntimeDesktopDepHits(file.content, rel);
    for (const hit of hits) {
      offenders.push(`${hit.label} :: ${hit.id}`);
    }
  }
  return offenders;
}

/**
 * @param {{ relativePath: string, content: string }[]} files
 */
export function assertNoRuntimeDesktopDepsInPackageFiles(files) {
  const offenders = collectRuntimeDesktopDepOffenders(files);
  if (offenders.length > 0) {
    throw new Error(
      `package contains runtime desktop/COM/.NET/Electron/child_process imports:\n- ${offenders.join("\n- ")}`,
    );
  }
}
