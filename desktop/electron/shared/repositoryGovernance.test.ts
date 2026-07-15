import { readFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

const repositoryRoot = path.resolve(process.cwd(), "..");

function readRepositoryFile(relativePath: string): string {
  return readFileSync(path.join(repositoryRoot, relativePath), "utf8");
}

describe("repository governance baseline", () => {
  it("provides a private vulnerability reporting path and response targets", () => {
    const securityPolicy = readRepositoryFile("SECURITY.md");

    expect(securityPolicy).toContain(
      "github.com/Laowang3366/excel-ai-plugin-new-feature/security/advisories/new",
    );
    expect(securityPolicy).toContain("2 个工作日内");
    expect(securityPolicy).toContain("Critical");
    expect(securityPolicy).not.toMatch(/TODO|example\.com|security@/iu);
  });

  it("assigns sensitive paths to the repository owner", () => {
    const codeowners = readRepositoryFile(".github/CODEOWNERS");

    expect(codeowners).toMatch(/^\* @Laowang3366$/mu);
    for (const sensitivePath of [
      "/.github/",
      "/desktop/electron/preload.ts",
      "/desktop/electron/agent/",
      "/desktop/dotnet/",
      "/product-site/src/",
      "/product-site/deploy/",
    ]) {
      expect(codeowners).toContain(`${sensitivePath} @Laowang3366`);
    }
  });

  it("documents required review and project-specific verification", () => {
    const contributing = readRepositoryFile("CONTRIBUTING.md");

    expect(contributing).toContain("至少两名人员批准");
    expect(contributing).toContain("branch protection/ruleset");
    expect(contributing).toContain("npm run office:test");
    expect(contributing).toContain("npm audit --audit-level=high");
  });

  it("keeps the data map factual about remote processing and deletion gaps", () => {
    const dataHandling = readRepositoryFile("docs/data-handling-and-privacy.md");

    expect(dataHandling).toContain("不是经法律审核的最终隐私政策");
    expect(dataHandling).toContain("远程数据处理\”默认关闭");
    expect(dataHandling).toContain("删除任一文件失败时不继续删除数据库投影");
    expect(dataHandling).toContain("本地数据隐私导出");
    expect(dataHandling).toContain("API Key、OCR Token 和自定义请求头秘密被省略或掩码");
    expect(dataHandling).toContain("ERASE LOCAL DATA");
    expect(dataHandling).toContain("只按固定白名单删除");
    expect(dataHandling).toContain("当前数据根路径会保留");
    expect(dataHandling).toContain("旧迁移目录");
    expect(dataHandling).toContain("可审计删除证明");
    expect(dataHandling).toContain("默认保留 90 天");
  });
});
