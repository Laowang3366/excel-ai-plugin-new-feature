import { describe, expect, it, vi } from "vitest";
import { DotNetOfficeDocumentBridge } from "./dotNetOfficeDocumentBridge";
import type { OfficeWorkerClient } from "./officeWorkerClient";

describe("DotNetOfficeDocumentBridge", () => {
  it("forwards document and object selection routes", async () => {
    const invoke = vi.fn().mockResolvedValue([]);
    const bridge = new DotNetOfficeDocumentBridge({ invoke } as unknown as OfficeWorkerClient);

    await bridge.listDocuments("excel");
    await bridge.activateDocument({ app: "word", filePath: "C:\\docs\\report.docx", instanceId: "word:1:2" });
    await bridge.listObjects({ app: "presentation", filePath: "C:\\slides.pptx", kind: "shape" });
    await bridge.activateObject({ app: "excel", filePath: "C:\\book.xlsx", locator: "sheet:Data" });

    expect(invoke).toHaveBeenNthCalledWith(1, "office.documents.list", { app: "excel" });
    expect(invoke).toHaveBeenNthCalledWith(2, "office.documents.activate", expect.objectContaining({ instanceId: "word:1:2" }));
    expect(invoke).toHaveBeenNthCalledWith(3, "office.objects.list", expect.objectContaining({ kind: "shape" }));
    expect(invoke).toHaveBeenNthCalledWith(4, "office.objects.activate", expect.objectContaining({ locator: "sheet:Data" }));
  });

  it("forwards transaction preparation and atomic restore routes", async () => {
    const invoke = vi.fn().mockResolvedValue([]);
    const bridge = new DotNetOfficeDocumentBridge({ invoke } as unknown as OfficeWorkerClient);
    const files = [{ filePath: "C:\\book.xlsx", existed: true, snapshotPath: "C:\\before.xlsx" }];

    await bridge.prepareTransaction(["C:\\book.xlsx"]);
    await bridge.restoreTransactionFiles(files);

    expect(invoke).toHaveBeenNthCalledWith(1, "office.transaction.prepare", { filePaths: ["C:\\book.xlsx"] });
    expect(invoke).toHaveBeenNthCalledWith(2, "office.transaction.restoreFiles", { files });
  });
});
