import { existsSync, readFileSync, readdirSync } from "node:fs";
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
  "docs/data-handling-and-privacy.md",
  "docs/codex-system-prompt-architecture.md",
  "docs/office-advanced-automation.md",
];

const retainedDocsFiles = [
  "README.md",
  "architecture-map.md",
  "development-standards.md",
  "update-and-release.md",
  "product-site-deployment.md",
  "data-handling-and-privacy.md",
  "codex-system-prompt-architecture.md",
  "office-advanced-automation.md",
  "enterprise-readiness-audit-2026-07-15.md",
];

const removedHistoricalNames = [
  "code-review-standards.md",
  "code-review-plan.md",
  "code-review-report",
  "dev-log.md",
  "docs/superpowers/",
  "sandbox-implementation-plan.md",
  "memory-gap-analysis.md",
  "session-2026-",
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

    expect(workflow).toContain('node-version: "22"');
    expect(workflow).toContain('dotnet-version: "8.0.422"');
    expect(workflow).not.toContain("test:coverage");
    expect(workflow).not.toContain("format:check");
  });

  it("keeps the incremental source-governance ratchet in desktop CI", () => {
    const workflow = readRepositoryFile(".github/workflows/ci.yml");
    const desktopPackage = JSON.parse(readRepositoryFile("desktop/package.json"));
    const developmentStandards = readRepositoryFile("docs/development-standards.md");

    expect(desktopPackage.scripts["governance:check"]).toBe(
      "node scripts/check-source-governance.cjs",
    );
    expect(workflow).toContain("npm run governance:check");
    expect(developmentStandards).toContain("普通功能 PR 禁止更新基线哈希");
  });

  it("indexes only retained current docs and excludes removed historical names", () => {
    const documentationIndex = readRepositoryFile("docs/README.md");
    const docsDir = path.join(repositoryRoot, "docs");
    const docsFiles = readdirSync(docsDir).filter((name) => name.endsWith(".md")).sort();

    expect(documentationIndex).toContain("已从仓库工作树移除");
    expect(documentationIndex).toContain("Git 历史");
    expect(docsFiles).toEqual([...retainedDocsFiles].sort());
    expect(existsSync(path.join(docsDir, "superpowers"))).toBe(false);

    for (const name of removedHistoricalNames) {
      expect(documentationIndex).not.toContain(name);
    }
  });
});
