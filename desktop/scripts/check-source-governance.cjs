const crypto = require("node:crypto");
const fs = require("node:fs");
const path = require("node:path");
const { spawnSync } = require("node:child_process");

const DESKTOP_ROOT = path.resolve(__dirname, "..");
const REPOSITORY_ROOT = path.resolve(DESKTOP_ROOT, "..");
const BASELINE_PATH = path.join(__dirname, "source-governance-baseline.json");
const PRETTIER_PATTERNS = [
  "src/**/*.{ts,tsx,css,json}",
  "electron/**/*.{ts,tsx,css,json}",
  "scripts/**/*.{ts,cjs,mjs,json}",
  "*.{ts,js,cjs,mjs,json,css}",
  "../product-site/src/**/*.{js,mjs,cjs,json,css}",
  "../product-site/scripts/**/*.{js,mjs,cjs,json}",
  "../product-site/test/**/*.{js,mjs,cjs,json}",
  "../product-site/*.{js,mjs,cjs,json,css}",
];
const SOURCE_ROOTS = [
  "desktop/src",
  "desktop/electron",
  "desktop/scripts",
  "product-site/src",
  "product-site/scripts",
];
const SOURCE_EXTENSIONS = new Set([".ts", ".tsx", ".css", ".mjs", ".cjs"]);

function normalizeRelativePath(value) {
  return value.replaceAll("\\", "/").replace(/^\.\//u, "");
}

function sha256File(filePath) {
  const content = fs.readFileSync(filePath, "utf8").replace(/\r\n/gu, "\n");
  return crypto.createHash("sha256").update(content, "utf8").digest("hex");
}

function countPhysicalLines(filePath) {
  const content = fs.readFileSync(filePath, "utf8").replace(/\r\n/gu, "\n");
  if (!content) return 0;
  return content.endsWith("\n")
    ? content.slice(0, -1).split("\n").length
    : content.split("\n").length;
}

function lineLimit(filePath) {
  const extension = path.extname(filePath).toLowerCase();
  if (extension === ".tsx") return 300;
  if (extension === ".css") return 500;
  return 400;
}

function walkFiles(directory) {
  if (!fs.existsSync(directory)) return [];
  return fs.readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const entryPath = path.join(directory, entry.name);
    return entry.isDirectory() ? walkFiles(entryPath) : [entryPath];
  });
}

function isOfficeSmokeScript(relativePath) {
  return /^desktop\/scripts\/smoke-[^/]+\.ts$/u.test(normalizeRelativePath(relativePath));
}

function collectProductionSources(repositoryRoot = REPOSITORY_ROOT, sourceRoots = SOURCE_ROOTS) {
  return sourceRoots
    .flatMap((root) => walkFiles(path.join(repositoryRoot, root)))
    .filter((filePath) => SOURCE_EXTENSIONS.has(path.extname(filePath).toLowerCase()))
    .filter((filePath) => !/\.(?:test|spec)\.[^.]+$/u.test(path.basename(filePath)));
}

function inspectSourceSizes({
  repositoryRoot = REPOSITORY_ROOT,
  sourceRoots = SOURCE_ROOTS,
  baseline = {},
} = {}) {
  const oversizedSources = collectProductionSources(repositoryRoot, sourceRoots)
    .map((filePath) => ({
      filePath,
      relativePath: normalizeRelativePath(path.relative(repositoryRoot, filePath)),
      lines: countPhysicalLines(filePath),
      limit: lineLimit(filePath),
      hash: sha256File(filePath),
    }))
    .filter((entry) => !isOfficeSmokeScript(entry.relativePath))
    .filter((entry) => entry.lines > entry.limit)
    .sort((left, right) => left.relativePath.localeCompare(right.relativePath));
  const violations = oversizedSources.filter(
    (entry) => baseline[entry.relativePath] !== entry.hash,
  );
  return { oversizedSources, violations };
}

function listPrettierDrift(desktopRoot = DESKTOP_ROOT) {
  const prettierBin = require.resolve("prettier/bin/prettier.cjs");
  const result = spawnSync(
    process.execPath,
    [prettierBin, "--list-different", ...PRETTIER_PATTERNS],
    { cwd: desktopRoot, encoding: "utf8", windowsHide: true },
  );
  if (result.status !== 0 && result.status !== 1) {
    throw new Error(result.stderr.trim() || `Prettier 检查异常退出: ${result.status}`);
  }
  return result.stdout
    .split(/\r?\n/gu)
    .map((filePath) => normalizeRelativePath(filePath.trim()))
    .filter(Boolean)
    .map((filePath) =>
      normalizeRelativePath(path.relative(REPOSITORY_ROOT, path.resolve(desktopRoot, filePath))),
    );
}

function inspectFormattingDrift({
  repositoryRoot = REPOSITORY_ROOT,
  driftFiles = listPrettierDrift(),
  baseline = {},
} = {}) {
  const violations = driftFiles.filter((relativePath) => {
    const filePath = path.join(repositoryRoot, relativePath);
    return !fs.existsSync(filePath) || baseline[relativePath] !== sha256File(filePath);
  });
  return { driftFiles, violations };
}

function sortedHashMap(entries) {
  return Object.fromEntries(
    [...entries]
      .sort((left, right) => left.relativePath.localeCompare(right.relativePath))
      .map((entry) => [entry.relativePath, entry.hash]),
  );
}

function writeBaseline() {
  const formattingEntries = listPrettierDrift().map((relativePath) => ({
    relativePath,
    hash: sha256File(path.join(REPOSITORY_ROOT, relativePath)),
  }));
  const { oversizedSources } = inspectSourceSizes();
  const baseline = {
    version: 1,
    generatedFrom: currentCommit(),
    formatting: sortedHashMap(formattingEntries),
    oversizedSources: sortedHashMap(oversizedSources),
  };
  fs.writeFileSync(BASELINE_PATH, `${JSON.stringify(baseline, null, 2)}\n`, "utf8");
  return baseline;
}

function currentCommit() {
  const result = spawnSync("git", ["rev-parse", "HEAD"], {
    cwd: REPOSITORY_ROOT,
    encoding: "utf8",
    windowsHide: true,
  });
  return result.status === 0 ? result.stdout.trim() : "unknown";
}

function run() {
  if (process.argv.includes("--write-baseline")) {
    const baseline = writeBaseline();
    console.log(
      `Source governance baseline written: formatting=${Object.keys(baseline.formatting).length}, oversized=${Object.keys(baseline.oversizedSources).length}`,
    );
    return;
  }
  const baseline = JSON.parse(fs.readFileSync(BASELINE_PATH, "utf8"));
  const formatting = inspectFormattingDrift({ baseline: baseline.formatting });
  const sizes = inspectSourceSizes({ baseline: baseline.oversizedSources });
  const messages = [
    ...formatting.violations.map((filePath) => `${filePath}: 新增或修改后仍不符合 Prettier`),
    ...sizes.violations.map(
      (entry) => `${entry.relativePath}: ${entry.lines} 行，超过 ${entry.limit} 行上限`,
    ),
  ];
  if (messages.length > 0) {
    throw new Error(`源码治理棘轮失败:\n${messages.join("\n")}`);
  }
  console.log(
    `Source governance passed: legacyFormatting=${formatting.driftFiles.length}, legacyOversized=${sizes.oversizedSources.length}`,
  );
}

if (require.main === module) {
  try {
    run();
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    process.exitCode = 1;
  }
}

module.exports = {
  countPhysicalLines,
  inspectFormattingDrift,
  inspectSourceSizes,
  lineLimit,
  normalizeRelativePath,
  sha256File,
};
