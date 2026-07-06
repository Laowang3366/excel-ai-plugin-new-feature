import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { afterEach, describe, expect, it } from "vitest";

import { listAuthorizedOfficeFiles } from "./ipcFileHandlers";
import { createPathAuthorizer } from "./ipcPathSecurity";

const tempDirs: string[] = [];

describe("ipcFileHandlers", () => {
  afterEach(async () => {
    await Promise.all(tempDirs.splice(0).map((dir) => fs.promises.rm(dir, { recursive: true, force: true })));
  });

  it("lists sorted Office files and authorizes returned file paths", async () => {
    const dir = await fs.promises.mkdtemp(path.join(os.tmpdir(), "office-files-"));
    tempDirs.push(dir);
    await fs.promises.writeFile(path.join(dir, "b.docx"), "doc");
    await fs.promises.writeFile(path.join(dir, "a.xlsx"), "book");
    await fs.promises.writeFile(path.join(dir, "slides.pptx"), "deck");
    await fs.promises.writeFile(path.join(dir, "note.txt"), "skip");

    const authorizer = createPathAuthorizer({
      getDataPath: () => "C:\\app\\data",
      getPinnedFolders: () => [],
      getExtraRoots: () => [],
    });

    const files = await listAuthorizedOfficeFiles(dir, authorizer);

    expect(files.map((file) => file.fileName)).toEqual(["a.xlsx", "b.docx", "slides.pptx"]);
    expect(files.every((file) => file.size > 0 && file.lastModified > 0)).toBe(true);
    expect(files.every((file) => authorizer.isAuthorizedPath(file.filePath))).toBe(true);
  });
});
