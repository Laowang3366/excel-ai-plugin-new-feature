import { describe, expect, it, vi } from "vitest";

import { createPathAuthorizer } from "../../main-modules/ipcPathSecurity";
import { reindexAuthorizedKnowledgeSources } from "./ipcKnowledgeHandlers";

function createKnowledgePathAuthorizer() {
  const authorizer = createPathAuthorizer({
    getDataPath: () => "C:\\app\\data",
    getPinnedFolders: () => [],
    getExtraRoots: () => [],
  });
  authorizer.authorizePath("C:\\picked\\allowed.pdf");
  authorizer.authorizePath("C:\\picked\\failing.pdf");
  return authorizer;
}

describe("reindexAuthorizedKnowledgeSources", () => {
  it("rechecks historical sources and skips paths outside current authorization", async () => {
    const indexFile = vi.fn(async (sourcePath: string) => ({
      sourcePath,
      success: true,
      entryCount: 1,
      durationMs: 1,
    }));
    const indexer = {
      listSources: () => [
        { sourcePath: "C:\\picked\\allowed.pdf", sourceName: "allowed.pdf" },
        { sourcePath: "C:\\secret\\denied.pdf", sourceName: "denied.pdf" },
      ],
      indexFile,
    };

    const results = await reindexAuthorizedKnowledgeSources(
      indexer as any,
      createKnowledgePathAuthorizer(),
    );

    expect(results).toHaveLength(2);
    expect(results[0].success).toBe(true);
    expect(results[1]).toMatchObject({
      sourcePath: "C:\\secret\\denied.pdf",
      success: false,
      entryCount: 0,
    });
    expect(indexFile).toHaveBeenCalledTimes(1);
  });

  it("continues reindexing after one authorized source fails", async () => {
    const indexFile = vi.fn(async (sourcePath: string) => {
      if (sourcePath.endsWith("failing.pdf")) throw new Error("parser failed");
      return {
        sourcePath,
        success: true,
        entryCount: 1,
        durationMs: 1,
      };
    });
    const indexer = {
      listSources: () => [
        { sourcePath: "C:\\picked\\failing.pdf", sourceName: "failing.pdf" },
        { sourcePath: "C:\\picked\\allowed.pdf", sourceName: "allowed.pdf" },
      ],
      indexFile,
    };

    const results = await reindexAuthorizedKnowledgeSources(
      indexer as any,
      createKnowledgePathAuthorizer(),
    );

    expect(results).toMatchObject([
      { sourcePath: "C:\\picked\\failing.pdf", success: false, error: "parser failed" },
      { sourcePath: "C:\\picked\\allowed.pdf", success: true },
    ]);
    expect(indexFile).toHaveBeenCalledTimes(2);
  });
});
