import type { PresentationBridge, WordDocumentBridge } from "../tools/contracts/office";
import { getOfficeWorkerClient, type OfficeWorkerClient } from "./officeWorkerClient";

interface DocumentConnectionStatus {
  connected: boolean;
  host: string;
  version?: string;
  documentName?: string;
  presentationName?: string;
}

export class DotNetWordBridge implements WordDocumentBridge {
  private status: DocumentConnectionStatus = { connected: false, host: "unknown" };

  constructor(private readonly client: OfficeWorkerClient = getOfficeWorkerClient()) {}

  isConnected(): boolean {
    return this.status.connected;
  }

  getHost(): string {
    return this.status.host;
  }

  async detectStatus(): Promise<DocumentConnectionStatus> {
    try {
      this.status = await this.client.invoke<DocumentConnectionStatus>("word.detectStatus");
    } catch {
      this.status = { connected: false, host: "unknown" };
    }
    return this.status;
  }

  async openDocument(
    filePath: string,
  ): Promise<{ success: boolean; documentName?: string; error?: string }> {
    try {
      const result = await this.client.invoke<{ success: boolean; documentName?: string }>(
        "word.open",
        { filePath },
      );
      if (result.success) await this.detectStatus();
      return result;
    } catch (error) {
      return { success: false, error: error instanceof Error ? error.message : String(error) };
    }
  }

  inspectDocument(): Promise<unknown> {
    return this.client.invoke("word.inspect");
  }

  readText(maxChars?: number): Promise<unknown> {
    return this.client.invoke("word.readText", { maxChars });
  }

  insertText(text: string, position?: string): Promise<unknown> {
    return this.client.invoke("word.insertText", { text, position });
  }

  insertHeading(text: string, level?: number, position?: string): Promise<unknown> {
    return this.client.invoke("word.insertHeading", { text, level, position });
  }

  replaceText(findText: string, replaceText: string, matchCase?: boolean): Promise<unknown> {
    return this.client.invoke("word.replaceText", { findText, replaceText, matchCase });
  }

  async saveDocument(saveAsPath?: string): Promise<{ success: boolean; error?: string }> {
    try {
      return await this.client.invoke("word.save", { saveAsPath });
    } catch (error) {
      return { success: false, error: error instanceof Error ? error.message : String(error) };
    }
  }
}

export class DotNetPresentationBridge implements PresentationBridge {
  private status: DocumentConnectionStatus = { connected: false, host: "unknown" };

  constructor(private readonly client: OfficeWorkerClient = getOfficeWorkerClient()) {}

  isConnected(): boolean {
    return this.status.connected;
  }

  getHost(): string {
    return this.status.host;
  }

  async detectStatus(): Promise<DocumentConnectionStatus> {
    try {
      this.status = await this.client.invoke<DocumentConnectionStatus>("presentation.detectStatus");
    } catch {
      this.status = { connected: false, host: "unknown" };
    }
    return this.status;
  }

  async openPresentation(
    filePath: string,
  ): Promise<{ success: boolean; presentationName?: string; error?: string }> {
    try {
      const result = await this.client.invoke<{ success: boolean; presentationName?: string }>(
        "presentation.open",
        { filePath },
      );
      if (result.success) await this.detectStatus();
      return result;
    } catch (error) {
      return { success: false, error: error instanceof Error ? error.message : String(error) };
    }
  }

  inspectPresentation(): Promise<unknown> {
    return this.client.invoke("presentation.inspect");
  }

  readSlide(slideIndex: number): Promise<unknown> {
    return this.client.invoke("presentation.readSlide", { slideIndex });
  }

  addSlide(title?: string, body?: string, layout?: string): Promise<unknown> {
    return this.client.invoke("presentation.addSlide", { title, body, layout });
  }

  setShapeText(
    slideIndex: number,
    text: string,
    shapeName?: string,
    shapeIndex?: number,
  ): Promise<unknown> {
    return this.client.invoke("presentation.setShapeText", {
      slideIndex,
      text,
      shapeName,
      shapeIndex,
    });
  }

  replaceText(findText: string, replaceText: string, matchCase?: boolean): Promise<unknown> {
    return this.client.invoke("presentation.replaceText", { findText, replaceText, matchCase });
  }

  async savePresentation(saveAsPath?: string): Promise<{ success: boolean; error?: string }> {
    try {
      return await this.client.invoke("presentation.save", { saveAsPath });
    } catch (error) {
      return { success: false, error: error instanceof Error ? error.message : String(error) };
    }
  }
}
