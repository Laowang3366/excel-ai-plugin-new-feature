import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { afterEach, describe, expect, it } from "vitest";

import type { KnowledgeEntry, KnowledgeSource } from "./types";
import { KnowledgeWriter } from "./knowledgeWriter";

describe("KnowledgeWriter", () => {
  const tempDirs: string[] = [];

  afterEach(() => {
    for (const dir of tempDirs.splice(0)) {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });

  it("writes model notes as searchable knowledge entries", async () => {
    const notesDir = fs.mkdtempSync(path.join(os.tmpdir(), "knowledge-writer-"));
    tempDirs.push(notesDir);
    const inserted: KnowledgeEntry[][] = [];
    const sources: KnowledgeSource[] = [];
    const writer = new KnowledgeWriter(
      {
        bulkInsert: (entries) => inserted.push(entries),
        upsertSource: (source) => sources.push(source),
      },
      {
        embedBatch: async (texts) => texts.map(() => [0.1, 0.2, 0.3]),
      },
      { notesDir }
    );

    const result = await writer.writeNote({
      title: "字段口径",
      content: "销售额字段必须扣除退款金额。",
      tags: ["销售", "口径"],
    });

    expect(result.entryCount).toBe(1);
    expect(result.sourceName).toMatch(/^note-\d{14}-[a-f0-9-]+/);
    expect(fs.existsSync(result.sourcePath)).toBe(true);
    expect(fs.readFileSync(result.sourcePath, "utf8")).toContain("销售额字段必须扣除退款金额。");
    expect(inserted[0][0]).toMatchObject({
      source: "note",
      sourceType: "md",
      sourcePath: result.sourcePath,
      metadata: expect.objectContaining({ title: "字段口径", tags: ["销售", "口径"] }),
      embedding: [0.1, 0.2, 0.3],
    });
    expect(sources[0]).toMatchObject({
      sourcePath: result.sourcePath,
      sourceType: "md",
      entryCount: 1,
    });
  });
});
