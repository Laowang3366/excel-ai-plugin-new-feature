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
  process.stderr.write("未找到 .NET 8 SDK。\n");
  process.exit(1);
}

const result = spawnSync(dotnet, [
  "list",
  "dotnet/Wengge.OfficeAutomation.sln",
  "package",
  "--vulnerable",
  "--include-transitive",
  "--format",
  "json",
], {
  cwd: path.resolve(__dirname, ".."),
  encoding: "utf8",
  windowsHide: true,
});
if (result.status !== 0) {
  process.stderr.write(result.stderr || result.stdout || "NuGet vulnerability scan failed\n");
  process.exit(result.status ?? 1);
}

const report = JSON.parse(result.stdout);
const vulnerable = (report.projects || []).flatMap(project =>
  (project.frameworks || []).flatMap(framework => [
    ...(framework.topLevelPackages || []),
    ...(framework.transitivePackages || []),
  ]).filter(pkg => Array.isArray(pkg.vulnerabilities) && pkg.vulnerabilities.length > 0)
    .map(pkg => ({ project: project.path, package: pkg.id, vulnerabilities: pkg.vulnerabilities })),
);
if (vulnerable.length > 0) {
  process.stderr.write(`${JSON.stringify(vulnerable, null, 2)}\n`);
  process.exit(1);
}
process.stdout.write("NuGet vulnerability scan passed.\n");
