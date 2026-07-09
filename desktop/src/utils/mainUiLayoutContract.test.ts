import * as fs from "fs";
import * as path from "path";
import { describe, expect, it } from "vitest";

const root = path.resolve(__dirname, "..");

function readSource(relativePath: string): string {
  return fs.readFileSync(path.join(root, relativePath), "utf-8");
}

describe("main UI layout contract", () => {
  it("exposes the redesigned titlebar controls", () => {
    const source = readSource("App.tsx");

    expect(source).toContain("titlebar-opacity-control");
    expect(source).toContain("titlebar-settings-btn");
  });

  it("keeps the sidebar search field visible in the expanded sidebar", () => {
    const source = readSource("components/sidebar/SidebarExpanded.tsx");

    expect(source).toContain("sidebar-search-field");
  });

  it("wraps the composer in the redesigned bottom shell", () => {
    const source = readSource("components/chat/ComposerArea.tsx");

    expect(source).toContain("composer-design-shell");
  });
});
