import { readFileSync, readdirSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";

const workflowsDirectory = resolve(process.cwd(), "..", ".github", "workflows");

function readWorkflow(name: string): string {
  return readFileSync(resolve(workflowsDirectory, name), "utf8");
}

describe("GitHub workflow provenance", () => {
  it("pins every external action to a full commit SHA", () => {
    const workflowFiles = readdirSync(workflowsDirectory).filter((name) => /\.ya?ml$/u.test(name));
    const mutableReferences: string[] = [];

    for (const name of workflowFiles) {
      const content = readWorkflow(name);
      for (const match of content.matchAll(/^\s*(?:-\s*)?uses:\s*([^\s#]+)/gmu)) {
        const reference = match[1];
        if (reference.startsWith("./") || reference.startsWith("docker://")) {
          continue;
        }
        if (!/^[^@\s]+@[0-9a-f]{40}$/u.test(reference)) {
          mutableReferences.push(`${name}: ${reference}`);
        }
      }
    }

    expect(mutableReferences).toEqual([]);
  });

  it("gates desktop releases and publishes a versioned SPDX SBOM", () => {
    const ci = readWorkflow("ci.yml");
    const release = readWorkflow("release-desktop.yml");

    expect(ci).toMatch(/^\s{2}workflow_call:\s*$/mu);
    expect(release).toContain("uses: ./.github/workflows/ci.yml");
    expect(release).toMatch(/build-and-sign:\s*\n\s+needs: quality-gates/u);
    expect(release).toContain("uses: anchore/sbom-action@");
    expect(release).toContain("syft-version: v1.46.0");
    expect(release).toContain(
      "Wengge-AI-Assistant-${{ steps.package-version.outputs.version }}.spdx.json",
    );
    expect(release).toContain("release-assets/desktop/release/Wengge-AI-Assistant-*.spdx.json");
    expect(release).toContain("if-no-files-found: error");
    expect(release).toContain("Verify release artifact inventory");
  });
});
