import { describe, expect, it } from "vitest";

import { getComposerPrimaryAction, isComposerSubmitKey } from "./ComposerArea";

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

describe("getComposerPrimaryAction", () => {
  it("shows stop while the model is running and the input is empty", () => {
    expect(getComposerPrimaryAction(true, false)).toBe("stop");
  });

  it("switches the same primary action to send when text is appended", () => {
    expect(getComposerPrimaryAction(true, true)).toBe("send");
  });

  it("uses the send action while idle", () => {
    expect(getComposerPrimaryAction(false, false)).toBe("send");
    expect(getComposerPrimaryAction(false, true)).toBe("send");
  });
});
