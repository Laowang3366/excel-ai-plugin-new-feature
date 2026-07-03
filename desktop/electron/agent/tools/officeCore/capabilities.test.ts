import { describe, expect, it } from "vitest";
import { findOfficeCapability } from "./capabilities";

describe("office capabilities", () => {
  it("declares first-stage document production capabilities", () => {
    expect(findOfficeCapability("excel", "insertChart")?.preferredEngine).toBe("openxml");
    expect(findOfficeCapability("word", "insertOrUpdateToc")?.fallback).toBe("needsCom");
    expect(findOfficeCapability("presentation", "deleteSlides")?.preferredEngine).toBe("openxml");
    expect(findOfficeCapability("presentation", "replacePictureSlot")?.writesFile).toBe(true);
  });

  it("declares visual snapshot as an Open XML first operation with COM fallback", () => {
    expect(findOfficeCapability("excel", "snapshot")?.fallback).toBe("needsCom");
    expect(findOfficeCapability("word", "snapshot")?.fallback).toBe("needsCom");
    expect(findOfficeCapability("presentation", "snapshot")?.fallback).toBe("needsCom");
  });
});
