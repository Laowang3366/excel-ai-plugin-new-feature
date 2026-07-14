const fs = require("node:fs");
const path = require("node:path");
const { createHash } = require("node:crypto");

const { zipSync } = require("fflate");

const ALLOWED_ROOTS = ["dist", "public/knowledge", "public/wps-jsa-bridge"];

function parseArgs(argv) {
  const result = { include: [] };
  for (let index = 0; index < argv.length; index += 1) {
    const key = argv[index];
    if (key === "--include") result.include.push(argv[++index]);
    else if (key.startsWith("--")) result[key.slice(2)] = argv[++index];
  }
  return result;
}

function walk(directory) {
  return fs.readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const target = path.join(directory, entry.name);
    return entry.isDirectory() ? walk(target) : [target];
  });
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (!args.id || !args["base-version"] || !args.sequence || !args["expires-at"] || !args.output) {
    throw new Error("用法: node scripts/create-hot-patch.cjs --id <id> --base-version <version> --sequence <n> --expires-at <ISO> --output <zip> [--include dist]");
  }
  const includes = args.include.length > 0 ? args.include : ALLOWED_ROOTS;
  for (const include of includes) {
    const normalized = include.replace(/\\/gu, "/").replace(/\/$/u, "");
    if (!ALLOWED_ROOTS.includes(normalized)) throw new Error(`不允许打包热补丁目录: ${include}`);
  }

  const files = {};
  const fileManifest = [];
  let fileCount = 0;
  for (const include of includes) {
    const root = path.resolve(include);
    if (!fs.existsSync(root)) continue;
    for (const filePath of walk(root)) {
      const relative = path.relative(process.cwd(), filePath).replace(/\\/gu, "/");
      const content = fs.readFileSync(filePath);
      files[relative] = content;
      fileManifest.push({
        path: relative,
        size: content.byteLength,
        sha256: createHash("sha256").update(content).digest("hex"),
      });
      fileCount += 1;
    }
  }
  if (fileCount === 0) throw new Error("热补丁没有可打包文件");
  const outputPath = path.resolve(args.output);
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, Buffer.from(zipSync(files, { level: 9 })));
  fs.writeFileSync(`${outputPath}.json`, `${JSON.stringify({
    id: args.id,
    baseVersion: args["base-version"],
    sequence: Number(args.sequence),
    publishedAt: new Date().toISOString(),
    expiresAt: new Date(args["expires-at"]).toISOString(),
    fileCount,
    files: fileManifest,
  }, null, 2)}\n`);
  console.log(JSON.stringify({ outputPath, fileCount }, null, 2));
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
