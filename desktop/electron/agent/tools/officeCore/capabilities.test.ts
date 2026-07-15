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

  it("declares deep Excel inspection and mutation operations with correct write boundaries", () => {
    for (const operation of [
      "inspectPowerQueries",
      "inspectCharts",
      "inspectWorkbookObjects",
      "captureWorkbookTemplate",
      "inspectWorkbookFormatting",
      "inspectPrintSettings",
      "inspectFormulaDependencies",
      "inspectFormulaBackups",
      "inspectFormulaProtection",
    ]) {
      expect(findOfficeCapability("excel", operation)).toMatchObject({
        preferredEngine: "com",
        writesFile: false,
      });
    }
    for (const operation of [
      "managePowerQuery",
      "formatChart",
      "manageWorkbookObject",
      "applyWorkbookTemplate",
      "configurePrint",
      "exportSheetsToPdf",
      "repairFormulaReferences",
      "convertFormulasToValues",
      "restoreFormulas",
      "manageFormulaProtection",
    ]) {
      expect(findOfficeCapability("excel", operation)).toMatchObject({
        preferredEngine: "com",
        writesFile: true,
      });
    }
  });

  it("declares advanced Word inspection, tracked editing, merge, and template operations", () => {
    for (const operation of [
      "inspectDocumentFormatting",
      "inspectReferences",
      "inspectRevisions",
      "inspectContentControls",
    ]) {
      expect(findOfficeCapability("word", operation)).toMatchObject({
        preferredEngine: "com",
        writesFile: false,
      });
    }
    for (const operation of [
      "formatLongDocument",
      "manageReferences",
      "manageRevisions",
      "compareDocuments",
      "applyTrackedChanges",
      "prepareMailMergeTemplate",
      "mailMerge",
      "batchMailMerge",
      "populateContentControls",
      "manageContentControls",
    ]) {
      expect(findOfficeCapability("word", operation)).toMatchObject({
        preferredEngine: "com",
        writesFile: true,
      });
    }
  });

  it("declares linked Office inspection and in-place refresh boundaries", () => {
    for (const app of ["word", "presentation"] as const) {
      expect(findOfficeCapability(app, "inspectLinkedOfficeContent")).toMatchObject({
        preferredEngine: "com",
        writesFile: false,
      });
      expect(findOfficeCapability(app, "refreshLinkedOfficeContent")).toMatchObject({
        preferredEngine: "com",
        writesFile: true,
      });
      expect(findOfficeCapability(app, "relinkLinkedOfficeContent")).toMatchObject({
        preferredEngine: "com",
        writesFile: true,
      });
    }
    for (const operation of [
      "exportRangeToWord",
      "exportRangeToPresentation",
      "buildReportPackage",
    ]) {
      expect(findOfficeCapability("excel", operation)).toMatchObject({
        preferredEngine: "com",
        writesFile: true,
      });
    }
  });
});
