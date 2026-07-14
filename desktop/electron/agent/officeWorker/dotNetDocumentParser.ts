import { getOfficeWorkerClient } from "./officeWorkerClient";

export interface ParsedOpenXmlChunk {
  content: string;
  sourceType: string;
  metadata: {
    sheetName?: string;
    tableRange?: string;
    headers?: string[];
    rowCount?: number;
    colCount?: number;
    slideNumber?: number;
    rows?: string[][];
  };
}

interface ParseDocumentResponse {
  filePath: string;
  chunks: ParsedOpenXmlChunk[];
}

export async function parseOpenXmlDocument(filePath: string): Promise<ParsedOpenXmlChunk[]> {
  const response = await getOfficeWorkerClient().invoke<ParseDocumentResponse>(
    "openxml.parseDocument",
    { filePath },
  );
  return response.chunks;
}
