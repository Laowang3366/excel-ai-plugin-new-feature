import { describe, expect, it } from "vitest";
import {
  formatKnowledgeFolderIndexSuccess,
  formatKnowledgeSourceStats,
  formatKnowledgeTime,
  getKnowledgeSourceTypeLabel,
  KNOWLEDGE_TEXT,
} from "./knowledgeSettingsText";

describe("knowledgeSettingsText", () => {
  it("formats indexed source counts", () => {
    expect(formatKnowledgeSourceStats(KNOWLEDGE_TEXT["en-US"], 3, 42)).toBe(
      "3 sources indexed, 42 knowledge entries",
    );
  });

  it("summarizes folder indexing results", () => {
    expect(
      formatKnowledgeFolderIndexSuccess(KNOWLEDGE_TEXT["en-US"], [
        { success: true, entryCount: 2 },
        { success: false, entryCount: 0 },
        { success: true, entryCount: 5 },
      ]),
    ).toBe("Folder indexed: 2 succeeded, 1 failed, 7 entries");
  });

  it("maps known source types and preserves unknown types", () => {
    expect(getKnowledgeSourceTypeLabel(KNOWLEDGE_TEXT["en-US"], "docx")).toBe("Word");
    expect(getKnowledgeSourceTypeLabel(KNOWLEDGE_TEXT["en-US"], "xlsx")).toBe("Workbook");
    expect(getKnowledgeSourceTypeLabel(KNOWLEDGE_TEXT["en-US"], "custom")).toBe("custom");
  });

  it("formats index timestamps with the selected locale", () => {
    expect(formatKnowledgeTime(Date.UTC(2026, 0, 2, 3, 4, 5), "en-US")).toContain("2026");
  });
});
