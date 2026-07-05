import os from "os";
import path from "path";
import { describe, expect, it } from "vitest";

import {
  assertAuthorizedPath,
  createPathAuthorizer,
} from "./ipcPathSecurity";

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

  it("allows temporary files without opening arbitrary system paths", () => {
    const tempFile = path.join(os.tmpdir(), "clipboard-123.png");
    const authorizer = createPathAuthorizer({
      getDataPath: () => "C:\\app\\data",
      getPinnedFolders: () => [],
      getExtraRoots: () => [],
    });

    expect(authorizer.isAuthorizedPath(tempFile)).toBe(true);
    expect(() => assertAuthorizedPath(authorizer, "C:\\secret\\token.txt")).toThrow("未授权访问路径");
  });
});
