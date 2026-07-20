import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import {
  PROMPT_MANIFEST,
  hasAdaptedPrompt,
  listPromptIds,
} from "../shared/prompts/loadPrompts";

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

  it("desktop-identical generated files match desktop SHA-256", () => {
    const identical = PROMPT_MANIFEST.files.filter(
      (f) => (f.mode ?? "desktop-identical") === "desktop-identical",
    );
    expect(identical.length).toBeGreaterThan(0);
    for (const entry of identical) {
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

  it("addin-adapted overlays exist for contaminated desktop scenarios", () => {
    const adaptedIds = [
      "system/base.zh-CN.md",
      "system/security.zh-CN.md",
      "scenarios/formula.zh-CN.md",
      "scenarios/office-tools.zh-CN.md",
      "scenarios/macro.zh-CN.md",
      "scenarios/general-office.zh-CN.md",
    ];
    for (const id of adaptedIds) {
      expect(hasAdaptedPrompt(id)).toBe(true);
      const entry = PROMPT_MANIFEST.files.find((f) => f.id === id);
      expect(entry?.mode).toBe("addin-adapted");
    }
    // Pure runtime reuse stays desktop-identical (no adapted override).
    expect(hasAdaptedPrompt("runtime/dynamic-array-enabled.zh-CN.md")).toBe(false);
    expect(hasAdaptedPrompt("runtime/environment.zh-CN.md")).toBe(false);
  });
});
