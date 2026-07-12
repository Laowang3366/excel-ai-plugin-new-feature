import { createReadStream, promises as fs } from "node:fs";
import path from "node:path";

function safeArtifactPath(releasesDir, artifact) {
  if (!artifact || path.basename(artifact) !== artifact) throw new Error("发布文件名无效");
  return path.join(releasesDir, artifact);
}

export function createReleaseStore(releasesDir) {
  async function readJson(fileName) {
    return JSON.parse(await fs.readFile(path.join(releasesDir, fileName), "utf8"));
  }

  return {
    readPublicRelease: () => readJson("release.json"),
    readSignedManifest: () => readJson("manifest.json"),
    async getInstaller() {
      const release = await readJson("release.json");
      const filePath = safeArtifactPath(releasesDir, release.installer.fileName);
      const stat = await fs.stat(filePath);
      return {
        release,
        filePath,
        fileName: release.installer.fileName,
        size: stat.size,
        stream: () => createReadStream(filePath),
      };
    },
  };
}
