import fs from "fs";
import os from "os";
import path from "path";
import { afterEach, describe, expect, it } from "vitest";

import {
  assertAuthorizedPath,
  createPathAuthorizer,
} from "./ipcPathSecurity";

const tempDirs: string[] = [];

afterEach(async () => {
  await Promise.all(
    tempDirs.splice(0).map((dir) => fs.promises.rm(dir, { recursive: true, force: true })),
  );
});

describe("ipcPathSecurity", () => {
  it("allows selected paths and their children while rejecting unrelated paths", () => {
    const authorizer = createPathAuthorizer({
      getDataPath: () => "C:\\app\\data",
      getPinnedFolders: () => ["C:\\work\\reports"],
      getExtraRoots: () => [],
    });

    authorizer.authorizePath("C:\\picked\\invoice.pdf");
    authorizer.authorizeRoot("C:\\picked-folder");

    expect(authorizer.isAuthorizedPath("C:\\picked\\invoice.pdf")).toBe(true);
    expect(authorizer.isAuthorizedPath("C:\\picked-folder\\nested\\deck.pptx")).toBe(true);
    expect(authorizer.isAuthorizedPath("C:\\work\\reports\\sales.xlsx")).toBe(true);
    expect(authorizer.isAuthorizedPath("C:\\Windows\\system.ini")).toBe(false);
  });

  it("trusts only the application-owned temp directory", async () => {
    const dataPath = await fs.promises.mkdtemp(path.join(os.tmpdir(), "path-auth-data-"));
    tempDirs.push(dataPath);
    const appTempFile = path.join(dataPath, "temp", "clipboard-123.png");
    const systemTempFile = path.join(os.tmpdir(), "untrusted-clipboard-123.png");
    const authorizer = createPathAuthorizer({
      getDataPath: () => dataPath,
      getPinnedFolders: () => [],
      getExtraRoots: () => [],
    });

    expect(authorizer.isAuthorizedPath(appTempFile)).toBe(true);
    expect(authorizer.isAuthorizedPath(systemTempFile)).toBe(false);
    expect(() => assertAuthorizedPath(authorizer, "C:\\secret\\token.txt")).toThrow("未授权访问路径");
  });

  it("rejects a junction that escapes an authorized root", async () => {
    const root = await fs.promises.mkdtemp(path.join(os.tmpdir(), "path-auth-root-"));
    const outside = await fs.promises.mkdtemp(path.join(os.tmpdir(), "path-auth-outside-"));
    tempDirs.push(root, outside);
    await fs.promises.writeFile(path.join(outside, "secret.txt"), "secret");
    const junction = path.join(root, "escape");
    await fs.promises.symlink(outside, junction, "junction");

    const authorizer = createPathAuthorizer({
      getDataPath: () => path.join(root, "app-data"),
      getPinnedFolders: () => [],
      getExtraRoots: () => [],
    });
    authorizer.authorizeRoot(root);

    const escapedPath = path.join(junction, "secret.txt");
    expect(authorizer.isAuthorizedPath(escapedPath)).toBe(false);
    expect(() => assertAuthorizedPath(authorizer, escapedPath)).toThrow("未授权访问路径");
  });
});
