import type {
  OfficeDocumentInfo,
  OfficeDocumentManagerBridge,
  OfficeObjectInfo,
} from "../tools/contracts/office";
import type { OfficeActionApp } from "../tools/officeCore/types";
import { getOfficeWorkerClient, type OfficeWorkerClient } from "./officeWorkerClient";

export class DotNetOfficeDocumentBridge implements OfficeDocumentManagerBridge {
  constructor(private readonly client: OfficeWorkerClient = getOfficeWorkerClient()) {}

  listDocuments(app?: OfficeActionApp): Promise<OfficeDocumentInfo[]> {
    return this.client.invoke("office.documents.list", { app });
  }

  activateDocument(input: {
    app: OfficeActionApp;
    filePath?: string;
    name?: string;
    index?: number;
    instanceId?: string;
  }): Promise<OfficeDocumentInfo> {
    return this.client.invoke("office.documents.activate", input);
  }

  listObjects(input: {
    app: OfficeActionApp;
    filePath: string;
    instanceId?: string;
    kind?: string;
  }): Promise<OfficeObjectInfo[]> {
    return this.client.invoke("office.objects.list", input);
  }

  activateObject(input: {
    app: OfficeActionApp;
    filePath: string;
    instanceId?: string;
    locator: string;
  }): Promise<OfficeObjectInfo> {
    return this.client.invoke("office.objects.activate", input);
  }

  prepareTransaction(filePaths: string[]): ReturnType<OfficeDocumentManagerBridge["prepareTransaction"]> {
    return this.client.invoke("office.transaction.prepare", { filePaths });
  }

  restoreTransactionFiles(
    files: Parameters<OfficeDocumentManagerBridge["restoreTransactionFiles"]>[0],
  ): ReturnType<OfficeDocumentManagerBridge["restoreTransactionFiles"]> {
    return this.client.invoke("office.transaction.restoreFiles", { files });
  }
}
