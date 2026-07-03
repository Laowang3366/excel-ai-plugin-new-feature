import { beforeEach, describe, expect, it } from "vitest";
import {
  getKnowledgeIndexer,
  getKnowledgeRetriever,
  getKnowledgeStore,
  getKnowledgeWriter,
  isKnowledgeAvailable,
  resetKnowledgeRegistry,
  setKnowledgeIndexer,
  setKnowledgeRetriever,
  setKnowledgeStore,
  setKnowledgeWriter,
} from "./knowledgeRegistry";

describe("knowledgeRegistry lifecycle", () => {
  beforeEach(() => {
    resetKnowledgeRegistry();
  });

  it("clears registered knowledge services when reset", () => {
    setKnowledgeRetriever({} as any);
    setKnowledgeStore({} as any);
    setKnowledgeIndexer({} as any);
    setKnowledgeWriter({} as any);

    expect(isKnowledgeAvailable()).toBe(true);

    resetKnowledgeRegistry();

    expect(getKnowledgeRetriever()).toBeNull();
    expect(getKnowledgeStore()).toBeNull();
    expect(getKnowledgeIndexer()).toBeNull();
    expect(getKnowledgeWriter()).toBeNull();
    expect(isKnowledgeAvailable()).toBe(false);
  });
});
