import path from "node:path";

import { createAnalyticsBackup, restoreAnalyticsBackup, verifyAnalyticsBackup } from "../src/analyticsBackup.mjs";

function parseArgs(argv) {
  const [command, ...tokens] = argv;
  const args = {};
  for (let index = 0; index < tokens.length; index += 1) {
    const token = tokens[index];
    if (!token.startsWith("--") || !tokens[index + 1]) throw new Error(`无效参数: ${token}`);
    args[token.slice(2)] = tokens[index + 1];
    index += 1;
  }
  return { command, args };
}

function requirePath(args, name, fallback) {
  const value = args[name] || fallback;
  if (!value) throw new Error(`缺少 --${name}`);
  return path.resolve(value);
}

const { command, args } = parseArgs(process.argv.slice(2));
let result;
if (command === "backup") {
  result = await createAnalyticsBackup({
    sourcePath: requirePath(args, "source", process.env.DATABASE_PATH),
    outputDir: requirePath(args, "output-dir", process.env.ANALYTICS_BACKUP_DIR),
    retain: args.retain === undefined ? 14 : Number(args.retain),
  });
} else if (command === "verify") {
  result = await verifyAnalyticsBackup(requirePath(args, "backup"));
} else if (command === "restore") {
  result = await restoreAnalyticsBackup({
    backupPath: requirePath(args, "backup"),
    targetPath: requirePath(args, "target"),
  });
} else {
  throw new Error("用法: analytics-backup.mjs <backup|verify|restore> [参数]");
}

console.log(JSON.stringify(result, null, 2));
