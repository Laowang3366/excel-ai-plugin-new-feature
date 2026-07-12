import { promises as fs } from "node:fs";
import path from "node:path";

import { sha256File, signManifest } from "../src/updateSigning.mjs";

function parseArgs(argv) {
  const result = {};
  for (let index = 0; index < argv.length; index += 1) {
    const key = argv[index];
    if (!key.startsWith("--")) continue;
    result[key.slice(2)] = argv[index + 1];
    index += 1;
  }
  return result;
}

function requireArg(args, name) {
  if (!args[name]) throw new Error(`缺少 --${name}`);
  return args[name];
}

async function copyArtifact(sourcePath, outputDir) {
  const fileName = path.basename(sourcePath);
  const destination = path.join(outputDir, fileName);
  await fs.copyFile(sourcePath, destination);
  const stat = await fs.stat(destination);
  return { fileName, destination, size: stat.size, sha256: await sha256File(destination) };
}

async function writeJsonAtomically(filePath, value) {
  const temporary = `${filePath}.tmp`;
  await fs.writeFile(temporary, `${JSON.stringify(value, null, 2)}\n`, "utf8");
  await fs.rename(temporary, filePath);
}

const args = parseArgs(process.argv.slice(2));
const version = requireArg(args, "version");
const installerPath = path.resolve(requireArg(args, "installer"));
const latestYmlPath = path.resolve(requireArg(args, "latest-yml"));
const blockmapPath = path.resolve(requireArg(args, "blockmap"));
const notesPath = path.resolve(requireArg(args, "notes-file"));
const privateKeyPath = path.resolve(requireArg(args, "private-key"));
const outputDir = path.resolve(requireArg(args, "output"));
const baseUrl = requireArg(args, "base-url").replace(/\/+$/u, "");

if (!/^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/u.test(version)) throw new Error("版本号格式无效");
const releaseNotes = JSON.parse(await fs.readFile(notesPath, "utf8"));
if (!Array.isArray(releaseNotes) || releaseNotes.some((note) => typeof note !== "string" || !note.trim())) {
  throw new Error("更新日志必须是非空字符串数组");
}

await fs.mkdir(outputDir, { recursive: true });
const installer = await copyArtifact(installerPath, outputDir);
await copyArtifact(blockmapPath, outputDir);
await fs.copyFile(latestYmlPath, path.join(outputDir, "latest.yml"));
const publishedAt = new Date().toISOString();

const unsignedManifest = {
  schemaVersion: 1,
  channel: "stable",
  version,
  publishedAt,
  releaseNotes,
  installer: {
    url: `${baseUrl}/releases/windows/${encodeURIComponent(installer.fileName)}`,
    sha256: installer.sha256,
    size: installer.size,
  },
};

if (args["hot-patch"]) {
  const patch = await copyArtifact(path.resolve(args["hot-patch"]), outputDir);
  unsignedManifest.hotPatch = {
    id: requireArg(args, "patch-id"),
    baseVersion: requireArg(args, "patch-base-version"),
    url: `${baseUrl}/releases/windows/${encodeURIComponent(patch.fileName)}`,
    sha256: patch.sha256,
    size: patch.size,
    restartRequired: true,
  };
}

const privateKey = await fs.readFile(privateKeyPath, "utf8");
const manifest = {
  ...unsignedManifest,
  signature: signManifest(unsignedManifest, privateKey),
};
const publicRelease = {
  version,
  publishedAt,
  releaseNotes,
  installer: {
    fileName: installer.fileName,
    size: installer.size,
    sha256: installer.sha256,
    downloadUrl: `${baseUrl}/download/windows`,
  },
};

await writeJsonAtomically(path.join(outputDir, "manifest.json"), manifest);
await writeJsonAtomically(path.join(outputDir, "release.json"), publicRelease);
console.log(JSON.stringify({ version, outputDir, installer: installer.fileName }, null, 2));
