import { promises as fs } from "node:fs";
import * as path from "node:path";

import { sha256File } from "../electron/main-modules/hotPatchManager";
import { verifyRemoteUpdateManifest } from "../electron/main-modules/updateManifest";

function readArg(name: string): string {
  const index = process.argv.indexOf(`--${name}`);
  const value = index >= 0 ? process.argv[index + 1] : undefined;
  if (!value) throw new Error(`缺少 --${name}`);
  return value;
}

const manifestPath = path.resolve(readArg("manifest"));
const publicKeyPath = path.resolve(readArg("public-key"));
const artifactDir = path.resolve(readArg("artifact-dir"));

const manifest = verifyRemoteUpdateManifest(
  JSON.parse(await fs.readFile(manifestPath, "utf8")),
  await fs.readFile(publicKeyPath, "utf8"),
);
if (!manifest.installer) throw new Error("更新清单缺少安装包");

const artifactName = path.basename(new URL(manifest.installer.url).pathname);
const artifactPath = path.join(artifactDir, decodeURIComponent(artifactName));
const stat = await fs.stat(artifactPath);
if (stat.size !== manifest.installer.size) throw new Error("安装包大小与更新清单不一致");

const sha256 = await sha256File(artifactPath);
if (sha256.toLowerCase() !== manifest.installer.sha256.toLowerCase()) {
  throw new Error("安装包 SHA-256 与更新清单不一致");
}

console.log(
  JSON.stringify(
    {
      version: manifest.version,
      artifact: path.basename(artifactPath),
      size: stat.size,
      sha256,
    },
    null,
    2,
  ),
);
