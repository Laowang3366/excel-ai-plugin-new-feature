import { readFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

const repositoryRoot = path.resolve(process.cwd(), "..");
const currentDocumentPaths = [
  "README.md",
  "overview.md",
  "docs/README.md",
  "docs/architecture-map.md",
  "docs/development-standards.md",
  "docs/update-and-release.md",
  "docs/product-site-deployment.md",
  "docs/codex-system-prompt-architecture.md",
  "docs/office-advanced-automation.md",
  "docs/code-review-standards.md",
];

function readRepositoryFile(relativePath: string): string {
  return readFileSync(path.join(repositoryRoot, relativePath), "utf8");
}

describe("current repository documentation", () => {
  it("uses the declared Node and .NET runtime baselines", () => {
    const desktopPackage = JSON.parse(readRepositoryFile("desktop/package.json"));
    const productSitePackage = JSON.parse(readRepositoryFile("product-site/package.json"));
    const globalJson = JSON.parse(readRepositoryFile("global.json"));
    const rootReadme = readRepositoryFile("README.md");

    expect(desktopPackage.engines.node).toBe(">=22.12.0");
    expect(productSitePackage.engines.node).toBe(">=22.12.0");
    expect(globalJson.sdk.version).toBe("8.0.422");
    expect(rootReadme).toContain("Node.js 22.12+");
    expect(rootReadme).toContain("global.json");
  });

  it("does not freeze volatile test or source line counts in current docs", () => {
    const currentDocuments = currentDocumentPaths.map(readRepositoryFile).join("\n");

    expect(currentDocuments).not.toMatch(/\b\d+\s*个测试文件/u);
    expect(currentDocuments).not.toMatch(/\b\d+\s*项测试/u);
    expect(currentDocuments).not.toMatch(/源码约\s*[\d,]+\s*个物理行/u);
  });

  it("does not describe unavailable tools as current CI gates", () => {
    const workflow = readRepositoryFile(".github/workflows/ci.yml");
    const reviewStandards = readRepositoryFile("docs/code-review-standards.md");

    expect(workflow).toContain('node-version: "22"');
    expect(workflow).toContain('dotnet-version: "8.0.422"');
    expect(workflow).not.toContain("test:coverage");
    expect(workflow).not.toContain("format:check");
    expect(reviewStandards).toContain("Coverage 候选（尚未启用）");
    expect(reviewStandards).toContain("Husky + lint-staged 候选（尚未启用）");
  });

  it("classifies changing activity logs as historical documentation", () => {
    const documentationIndex = readRepositoryFile("docs/README.md");
    expect(documentationIndex).toContain("`dev-log.md`");
    expect(documentationIndex).toContain("不作为当前运行能力或发布流程的依据");
  });
});
