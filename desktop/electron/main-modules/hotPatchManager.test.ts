import { createHash } from "node:crypto";
import { promises as fs } from "node:fs";
import * as os from "node:os";
import * as path from "node:path";

import { strToU8, zipSync } from "fflate";
import { afterEach, describe, expect, it } from "vitest";

import {
  activateInstalledHotPatch,
  installHotPatchArchive,
  isAllowedHotPatchPath,
  resolveHotPatchPath,
} from "./hotPatchManager";

const temporaryDirectories: string[] = [];

afterEach(async () => {
  delete process.env.WENGE_HOT_PATCH_ROOT;
  await Promise.all(temporaryDirectories.splice(0).map((directory) => fs.rm(directory, { recursive: true, force: true })));
});

async function createArchive(files: Record<string, string>) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), "wenge-patch-"));
  temporaryDirectories.push(root);
  const buffer = Buffer.from(zipSync(Object.fromEntries(
    Object.entries(files).map(([name, content]) => [name, strToU8(content)]),
  )));
  const archivePath = path.join(root, "patch.zip");
  await fs.writeFile(archivePath, buffer);
  return {
    root,
    archivePath,
    size: buffer.byteLength,
    sha256: createHash("sha256").update(buffer).digest("hex"),
  };
}

describe("hot patch paths", () => {
  it("allows renderer and selected public resources only", () => {
    expect(isAllowedHotPatchPath("dist/index.html")).toBe(true);
    expect(isAllowedHotPatchPath("public/knowledge/builtin-knowledge.json")).toBe(true);
    expect(isAllowedHotPatchPath("electron/main.js")).toBe(false);
    expect(isAllowedHotPatchPath("../outside.txt")).toBe(false);
  });
});

describe("installHotPatchArchive", () => {
  it("installs and activates a verified patch atomically", async () => {
    const archive = await createArchive({
      "dist/index.html": "<main>patched</main>",
      "dist/assets/app.js": "console.log('patched')",
    });
    const userDataPath = path.join(archive.root, "user-data");
    const state = await installHotPatchArchive({
      archivePath: archive.archivePath,
      currentVersion: "0.1.79",
      userDataPath,
      descriptor: {
        id: "patch-001",
        baseVersion: "0.1.79",
        url: "https://plugin.shelelove.top/releases/patches/patch-001.zip",
        sha256: archive.sha256,
        size: archive.size,
        restartRequired: true,
      },
    });

    expect(await fs.readFile(path.join(state.rootPath, "dist/index.html"), "utf8"))
      .toContain("patched");
    expect(activateInstalledHotPatch("0.1.79", userDataPath)).toBe("patch-001");
    expect(resolveHotPatchPath("dist/index.html")).toBe(path.join(state.rootPath, "dist/index.html"));
  });

  it("rejects files outside the patch allowlist", async () => {
    const archive = await createArchive({ "electron/main.js": "unsafe" });
    await expect(installHotPatchArchive({
      archivePath: archive.archivePath,
      currentVersion: "0.1.79",
      userDataPath: path.join(archive.root, "user-data"),
      descriptor: {
        id: "patch-unsafe",
        baseVersion: "0.1.79",
        url: "https://plugin.shelelove.top/releases/patches/patch-unsafe.zip",
        sha256: archive.sha256,
        size: archive.size,
        restartRequired: true,
      },
    })).rejects.toThrow("不允许的文件");
  });
});
