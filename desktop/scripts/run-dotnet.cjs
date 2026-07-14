const { existsSync } = require("node:fs");
const path = require("node:path");
const { spawnSync } = require("node:child_process");

const executable = process.platform === "win32" ? "dotnet.exe" : "dotnet";
const candidates = [
  process.env.DOTNET_ROOT && path.join(process.env.DOTNET_ROOT, executable),
  process.env.LOCALAPPDATA && path.join(process.env.LOCALAPPDATA, "Microsoft", "dotnet", executable),
  process.env.ProgramFiles && path.join(process.env.ProgramFiles, "dotnet", executable),
  executable,
].filter(Boolean);
const dotnet = candidates.find(candidate => candidate === executable || existsSync(candidate));
if (!dotnet) {
  process.stderr.write("未找到 .NET 8 SDK，请先安装 Microsoft .NET 8 SDK。\n");
  process.exit(1);
}

const result = spawnSync(dotnet, process.argv.slice(2), {
  cwd: path.resolve(__dirname, ".."),
  stdio: "inherit",
  windowsHide: true,
});
if (result.error) {
  process.stderr.write(`${result.error.message}\n`);
  process.exit(1);
}
process.exit(result.status ?? 1);
