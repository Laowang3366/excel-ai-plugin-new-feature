import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { PROMPT_MANIFEST, listPromptIds } from "../shared/prompts/loadPrompts";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const repoRoot = path.resolve(root, "..");

describe("prompt sync manifest", () => {
  it("tracks Excel-related sources only", () => {
    const ids = listPromptIds();
    expect(ids).toContain("system/base.zh-CN.md");
    expect(ids).toContain("scenarios/formula.zh-CN.md");
    expect(ids).toContain("runtime/dynamic-array-enabled.zh-CN.md");
    expect(ids.some((id) => id.includes("ocr"))).toBe(false);
  });

  it("matches source SHA-256 to prevent drift", () => {
    expect(PROMPT_MANIFEST.files.length).toBeGreaterThan(0);
    for (const entry of PROMPT_MANIFEST.files) {
      const sourceAbs = path.join(repoRoot, entry.sourcePath);
      const generatedAbs = path.join(root, entry.generatedPath);
      const source = readFileSync(sourceAbs);
      const generated = readFileSync(generatedAbs);
      const sourceHash = createHash("sha256").update(source).digest("hex");
      const generatedHash = createHash("sha256").update(generated).digest("hex");
      expect(sourceHash).toBe(entry.sha256);
      expect(generatedHash).toBe(entry.sha256);
    }
  });
});
