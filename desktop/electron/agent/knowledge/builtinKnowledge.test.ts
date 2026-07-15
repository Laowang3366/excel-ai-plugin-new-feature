import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { afterEach, describe, expect, it, vi } from "vitest";

import { indexBuiltinKnowledge } from "./builtinKnowledge";

let tempRoots: string[] = [];

afterEach(() => {
  vi.restoreAllMocks();
  for (const root of tempRoots) fs.rmSync(root, { recursive: true, force: true });
  tempRoots = [];
});

describe("indexBuiltinKnowledge", () => {
  it("does not bundle a second copy of the formula methodology", () => {
    const knowledgeRoot = path.resolve(__dirname, "../../../public/knowledge");
    const manifest = JSON.parse(
      fs.readFileSync(path.join(knowledgeRoot, "builtin-knowledge.json"), "utf-8"),
    ) as { version: number; files: unknown[] };

    expect(manifest).toEqual({ version: 3, files: [] });
    expect(fs.readdirSync(knowledgeRoot).sort()).toEqual(["builtin-knowledge.json"]);
  });

  it("removes retired builtin sources when the manifest is empty", async () => {
    const root = createBuiltinRoot({ files: [] });
    const indexer = {
      indexFile: vi.fn(),
      deleteSource: vi.fn(async () => undefined),
      listSources: vi.fn(() => [
        {
          sourcePath:
            "C:\\old-app\\public\\knowledge\\excel-wps-formula-problem-solving-methodology.md",
          sourceName: "excel-wps-formula-problem-solving-methodology.md",
        },
        {
          sourcePath: "D:\\user-documents\\notes.md",
          sourceName: "notes.md",
        },
        {
          sourcePath: "D:\\user-project\\public\\knowledge\\custom-rules.md",
          sourceName: "custom-rules.md",
        },
      ]),
    };

    vi.spyOn(process, "cwd").mockReturnValue(root);
    const results = await indexBuiltinKnowledge(indexer as any);

    expect(results).toEqual([]);
    expect(indexer.indexFile).not.toHaveBeenCalled();
    expect(indexer.deleteSource).toHaveBeenCalledOnce();
    expect(indexer.deleteSource).toHaveBeenCalledWith(
      "C:\\old-app\\public\\knowledge\\excel-wps-formula-problem-solving-methodology.md",
    );
  });

  it("keeps existing builtin sources when the manifest is unavailable", async () => {
    const root = createBuiltinRoot({ manifest: false });
    const indexer = {
      indexFile: vi.fn(),
      deleteSource: vi.fn(),
      listSources: vi.fn(),
    };

    vi.spyOn(process, "cwd").mockReturnValue(root);
    const results = await indexBuiltinKnowledge(indexer as any);

    expect(results).toEqual([]);
    expect(indexer.listSources).not.toHaveBeenCalled();
    expect(indexer.deleteSource).not.toHaveBeenCalled();
  });
});

function createBuiltinRoot(input: { manifest?: boolean; files?: unknown[] }): string {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "builtin-knowledge-"));
  tempRoots.push(root);
  const knowledgeRoot = path.join(root, "public", "knowledge");
  fs.mkdirSync(knowledgeRoot, { recursive: true });
  if (input.manifest !== false) {
    fs.writeFileSync(
      path.join(knowledgeRoot, "builtin-knowledge.json"),
      JSON.stringify({ version: 3, files: input.files ?? [] }),
    );
  }
  return root;
}
