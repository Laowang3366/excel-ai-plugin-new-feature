const { spawnSync } = require("child_process");
const path = require("path");

const env = {
  ...process.env,
  ELECTRON_MIRROR:
    process.env.ELECTRON_MIRROR || "https://npmmirror.com/mirrors/electron/",
  ELECTRON_BUILDER_BINARIES_MIRROR:
    process.env.ELECTRON_BUILDER_BINARIES_MIRROR ||
    "https://npmmirror.com/mirrors/electron-builder-binaries/",
};

const rebuildCliPath = path.join(path.dirname(require.resolve("@electron/rebuild")), "cli.js");
const rebuildResult = spawnSync(
  process.execPath,
  [rebuildCliPath, "-f", "-w", "better-sqlite3"],
  {
    stdio: "inherit",
    env,
  }
);

if (rebuildResult.error) {
  console.error(rebuildResult.error);
  process.exit(1);
}
if (rebuildResult.status !== 0) {
  process.exit(rebuildResult.status ?? 1);
}

const cliPath = require.resolve("electron-builder/cli.js");
const result = spawnSync(process.execPath, [cliPath, ...process.argv.slice(2)], {
  stdio: "inherit",
  env,
});

if (result.error) {
  console.error(result.error);
  process.exit(1);
}

process.exit(result.status ?? 1);
