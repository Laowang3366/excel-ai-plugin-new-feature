/**
 * Shared CLI argument parsing for WPS install tools.
 */

function isOptionToken(t) {
  return typeof t === "string" && t.startsWith("-");
}

function takeValue(argv, i, flag) {
  const v = argv[i + 1];
  if (v == null || isOptionToken(v)) {
    throw new Error(`${flag} requires a value`);
  }
  return { value: v, next: i + 1 };
}

/**
 * @param {string[]} argv
 * @param {{ allowGitSha?: boolean, allowPackageDir?: boolean }} flags
 */
export function parseWpsInstallCliArgs(argv, flags = {}) {
  const allowGitSha = flags.allowGitSha === true;
  const allowPackageDir = flags.allowPackageDir === true;
  const out = {
    gitSha: null,
    packageDir: null,
    appData: null,
    help: false,
  };
  for (let i = 0; i < argv.length; i += 1) {
    const a = argv[i];
    if (a === "-h" || a === "--help") {
      out.help = true;
      continue;
    }
    if (a === "--app-data") {
      const t = takeValue(argv, i, "--app-data");
      out.appData = t.value;
      i = t.next;
      continue;
    }
    if (allowGitSha && a === "--git-sha") {
      const t = takeValue(argv, i, "--git-sha");
      out.gitSha = t.value;
      i = t.next;
      continue;
    }
    if (allowPackageDir && a === "--package-dir") {
      const t = takeValue(argv, i, "--package-dir");
      out.packageDir = t.value;
      i = t.next;
      continue;
    }
    throw new Error(`Unknown arg: ${a}`);
  }
  if (out.packageDir && out.gitSha) {
    throw new Error("--package-dir installs an existing package; do not also pass --git-sha");
  }
  return out;
}
