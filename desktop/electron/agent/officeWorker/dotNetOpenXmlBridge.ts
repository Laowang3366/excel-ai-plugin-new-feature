import type { OfficeFileBridge } from "../tools/contracts/office";
import type { OfficeActionInput, OfficeActionResult } from "../tools/officeCore/types";
import { actionTimeout } from "./dotNetOfficeActionBridge";
import { getOfficeWorkerClient, type OfficeWorkerClient } from "./officeWorkerClient";

export class DotNetOpenXmlBridge implements OfficeFileBridge {
  constructor(private readonly client: OfficeWorkerClient = getOfficeWorkerClient()) {}

  executeAction(input: OfficeActionInput): Promise<OfficeActionResult> {
    return this.client.invoke("openxml.action.execute", { ...input }, actionTimeout(input));
  }

  inspectFile(filePath: string): Promise<unknown> {
    return this.client.invoke("openxml.inspect", { filePath });
  }

  replaceText(input: Parameters<OfficeFileBridge["replaceText"]>[0]): Promise<unknown> {
    return this.client.invoke("openxml.replaceText", input);
  }

  inspectLayout(input: Parameters<OfficeFileBridge["inspectLayout"]>[0]): Promise<unknown> {
    return this.client.invoke("openxml.inspectLayout", input);
  }

  inspectTable(input: Parameters<OfficeFileBridge["inspectTable"]>[0]): Promise<unknown> {
    return this.client.invoke("openxml.inspectTable", input);
  }

  applyTableStyle(input: Parameters<OfficeFileBridge["applyTableStyle"]>[0]): Promise<unknown> {
    return this.client.invoke("openxml.applyTableStyle", input);
  }

  snapshot(input: Parameters<OfficeFileBridge["snapshot"]>[0]): Promise<unknown> {
    return this.client.invoke("openxml.snapshot", input);
  }
}
