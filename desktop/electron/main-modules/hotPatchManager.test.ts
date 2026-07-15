import { createHash } from "node:crypto";
import { promises as fs } from "node:fs";
import * as os from "node:os";
import * as path from "node:path";

import { strToU8, zipSync } from "fflate";
import { afterEach, describe, expect, it } from "vitest";

import {
  acknowledgeActiveHotPatchHealth,
  activateInstalledHotPatch,
  applyHotPatchPolicy,
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
    files: Object.entries(files).map(([filePath, content]) => ({
      path: filePath,
      size: Buffer.byteLength(content),
      sha256: createHash("sha256").update(content).digest("hex"),
    })),
  };
}

function descriptor(archive: Awaited<ReturnType<typeof createArchive>>, id = "patch-001", sequence = 1) {
  return {
    id,
    baseVersion: "0.1.79",
    sequence,
    publishedAt: new Date(Date.now() - 60_000).toISOString(),
    expiresAt: new Date(Date.now() + 24 * 60 * 60 * 1000).toISOString(),
    url: `https://plugin.shelelove.top/releases/patches/${id}.zip`,
    sha256: archive.sha256,
    size: archive.size,
    files: archive.files,
    restartRequired: true as const,
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
      descriptor: descriptor(archive),
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
      descriptor: descriptor(archive, "patch-unsafe"),
    })).rejects.toThrow("不允许的文件");
  });

  it("rejects replayed or lower patch sequences", async () => {
    const first = await createArchive({ "dist/index.html": "first" });
    const userDataPath = path.join(first.root, "user-data");
    await installHotPatchArchive({
      archivePath: first.archivePath,
      currentVersion: "0.1.79",
      userDataPath,
      descriptor: descriptor(first, "patch-002", 2),
    });
    const replay = await createArchive({ "dist/index.html": "replay" });
    await expect(installHotPatchArchive({
      archivePath: replay.archivePath,
      currentVersion: "0.1.79",
      userDataPath,
      descriptor: descriptor(replay, "patch-001", 1),
    })).rejects.toThrow("低于安全基线");
  });

  it("refuses to activate a patch whose installed files were modified", async () => {
    const archive = await createArchive({ "dist/index.html": "verified" });
    const userDataPath = path.join(archive.root, "user-data");
    const state = await installHotPatchArchive({
      archivePath: archive.archivePath,
      currentVersion: "0.1.79",
      userDataPath,
      descriptor: descriptor(archive),
    });
    await fs.writeFile(path.join(state.rootPath, "dist/index.html"), "tampered");

    expect(activateInstalledHotPatch("0.1.79", userDataPath)).toBeNull();
  });

  it("rolls back a patch on the next startup when renderer health was never acknowledged", async () => {
    const archive = await createArchive({ "dist/index.html": "health-check" });
    const userDataPath = path.join(archive.root, "user-data");
    await installHotPatchArchive({
      archivePath: archive.archivePath,
      currentVersion: "0.1.79",
      userDataPath,
      descriptor: descriptor(archive),
    });

    expect(activateInstalledHotPatch("0.1.79", userDataPath)).toBe("patch-001");
    delete process.env.WENGE_HOT_PATCH_ROOT;
    expect(activateInstalledHotPatch("0.1.79", userDataPath)).toBeNull();
  });

  it("keeps an acknowledged renderer patch active on later startups", async () => {
    const archive = await createArchive({ "dist/index.html": "healthy" });
    const userDataPath = path.join(archive.root, "user-data");
    await installHotPatchArchive({
      archivePath: archive.archivePath,
      currentVersion: "0.1.79",
      userDataPath,
      descriptor: descriptor(archive),
    });

    expect(activateInstalledHotPatch("0.1.79", userDataPath)).toBe("patch-001");
    expect(await acknowledgeActiveHotPatchHealth(userDataPath)).toBe(true);
    delete process.env.WENGE_HOT_PATCH_ROOT;
    expect(activateInstalledHotPatch("0.1.79", userDataPath)).toBe("patch-001");
  });

  it("disables an active patch when a signed policy revokes its id", async () => {
    const archive = await createArchive({ "dist/index.html": "revoked" });
    const userDataPath = path.join(archive.root, "user-data");
    await installHotPatchArchive({
      archivePath: archive.archivePath,
      currentVersion: "0.1.79",
      userDataPath,
      descriptor: descriptor(archive),
    });
    expect(activateInstalledHotPatch("0.1.79", userDataPath)).toBe("patch-001");

    const disabled = await applyHotPatchPolicy(userDataPath, "0.1.79", {
      revokedPatchIds: ["patch-001"],
      minimumSafeSequenceByBaseVersion: {},
    });

    expect(disabled).toBe(true);
    expect(process.env.WENGE_HOT_PATCH_ROOT).toBeUndefined();
    expect(activateInstalledHotPatch("0.1.79", userDataPath)).toBeNull();
  });

  it("rejects installation below the signed minimum safe sequence", async () => {
    const archive = await createArchive({ "dist/index.html": "old" });
    const userDataPath = path.join(archive.root, "user-data");
    await applyHotPatchPolicy(userDataPath, "0.1.79", {
      revokedPatchIds: [],
      minimumSafeSequenceByBaseVersion: { "0.1.79": 3 },
    });

    await expect(installHotPatchArchive({
      archivePath: archive.archivePath,
      currentVersion: "0.1.79",
      userDataPath,
      descriptor: descriptor(archive, "patch-002", 2),
    })).rejects.toThrow("远程安全基线");
  });
});
