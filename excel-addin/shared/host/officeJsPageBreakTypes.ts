/** ExcelApi 1.9 Worksheet page break collection surface (manual only). */
import type { ExcelRange } from "./officeJsRuntime";

export interface ExcelPageBreak {
  getCellAfterBreak(): ExcelRange;
  load(props: string): void;
}

export interface ExcelPageBreakCollection {
  items: ExcelPageBreak[];
  add(reference: string | ExcelRange): ExcelPageBreak;
  removePageBreaks(): void;
  load(props: string): void;
}
