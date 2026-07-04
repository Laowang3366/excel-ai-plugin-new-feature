import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

const declarationText = readFileSync(new URL("./electronApi.d.ts", import.meta.url), "utf8");

describe("electronApi type declarations", () => {
  it("keeps legacy AttachedFile as an alias of FileAttachment", () => {
    expect(declarationText).toContain("export interface FileAttachment");
    expect(declarationText).toContain("export type AttachedFile = FileAttachment;");
    expect(declarationText).not.toMatch(/export\s+interface\s+AttachedFile\b/);
  });
});
