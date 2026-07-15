import fs from "node:fs";
import os from "node:os";
import path from "node:path";

import { afterEach, describe, expect, it } from "vitest";

const { inspectFormattingDrift, inspectSourceSizes, sha256File } =
  require("../../scripts/check-source-governance.cjs") as {
    inspectFormattingDrift: (options: {
      repositoryRoot: string;
      driftFiles: string[];
      baseline: Record<string, string>;
    }) => { violations: string[] };
    inspectSourceSizes: (options: {
      repositoryRoot: string;
      sourceRoots: string[];
      baseline: Record<string, string>;
    }) => { violations: Array<{ relativePath: string; lines: number; limit: number }> };
    sha256File: (filePath: string) => string;
  };

const temporaryDirectories: string[] = [];

afterEach(() => {
  for (const directory of temporaryDirectories.splice(0)) {
    fs.rmSync(directory, { recursive: true, force: true });
  }
});

describe("source governance ratchet", () => {
  it("normalizes line endings before hashing the baseline", () => {
    const root = makeRepository();
    const lfPath = writeSource(root, "desktop/src/lf.ts", "const value = 1;\n");
    const crlfPath = writeSource(root, "desktop/src/crlf.ts", "const value = 1;\r\n");

    expect(sha256File(lfPath)).toBe(sha256File(crlfPath));
  });

  it("allows only byte-identical legacy formatting drift", () => {
    const root = makeRepository();
    const relativePath = "desktop/src/legacy.ts";
    const filePath = writeSource(root, relativePath, "const value=1\n");
    const baseline = { [relativePath]: sha256File(filePath) };

    expect(
      inspectFormattingDrift({
        repositoryRoot: root,
        driftFiles: [relativePath],
        baseline,
      }).violations,
    ).toEqual([]);

    fs.appendFileSync(filePath, "const next=2\n");
    expect(
      inspectFormattingDrift({
        repositoryRoot: root,
        driftFiles: [relativePath],
        baseline,
      }).violations,
    ).toEqual([relativePath]);
  });

  it("rejects new or modified oversized production files", () => {
    const root = makeRepository();
    const relativePath = "desktop/src/legacy.ts";
    const filePath = writeSource(root, relativePath, `${"// line\n".repeat(400)}export {};\n`);
    const baseline = { [relativePath]: sha256File(filePath) };

    expect(
      inspectSourceSizes({
        repositoryRoot: root,
        sourceRoots: ["desktop/src"],
        baseline,
      }).violations,
    ).toEqual([]);

    fs.appendFileSync(filePath, "// growth\n");
    expect(
      inspectSourceSizes({
        repositoryRoot: root,
        sourceRoots: ["desktop/src"],
        baseline,
      }).violations,
    ).toEqual([expect.objectContaining({ relativePath, lines: 402, limit: 400 })]);
  });

  it("accepts a formerly oversized source after it is split below the limit", () => {
    const root = makeRepository();
    const relativePath = "desktop/src/legacy.ts";
    writeSource(root, relativePath, "export {};\n");

    expect(
      inspectSourceSizes({
        repositoryRoot: root,
        sourceRoots: ["desktop/src"],
        baseline: { [relativePath]: "old-hash" },
      }).violations,
    ).toEqual([]);
  });
});

function makeRepository(): string {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "source-governance-"));
  temporaryDirectories.push(root);
  return root;
}

function writeSource(root: string, relativePath: string, content: string): string {
  const filePath = path.join(root, relativePath);
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, content, "utf8");
  return filePath;
}
