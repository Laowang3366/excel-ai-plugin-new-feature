import { describe, expect, it } from "vitest";

import { isComposerSubmitKey } from "./ComposerArea";

describe("isComposerSubmitKey", () => {
  it("submits on Enter without Shift", () => {
    expect(isComposerSubmitKey("Enter", false)).toBe(true);
  });

  it("does not submit on Shift+Enter", () => {
    expect(isComposerSubmitKey("Enter", true)).toBe(false);
  });

  it("does not submit Enter while an IME composition is active", () => {
    expect(isComposerSubmitKey("Enter", false, true)).toBe(false);
  });

  it("does not submit for other keys", () => {
    expect(isComposerSubmitKey("Space", false)).toBe(false);
  });
});
