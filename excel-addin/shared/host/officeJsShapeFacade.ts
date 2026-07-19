/** Official ExcelApi 1.9 Shape facade (extracted for ≤400 line budget). */

export interface ExcelTextFrame {
  hasText: boolean;
  textRange: { text: string; load(props: string): void };
  load(props: string): void;
}

export interface ExcelShape {
  name: string;
  type: string;
  geometricShapeType: string | null;
  left: number;
  top: number;
  width: number;
  height: number;
  visible: boolean;
  textFrame: ExcelTextFrame;
  isNullObject?: boolean;
  delete(): void;
  load(props: string): void;
}

export interface ExcelShapeCollection {
  items: ExcelShape[];
  addGeometricShape(geometricShapeType: string): ExcelShape;
  addTextBox(text?: string): ExcelShape;
  getItem(key: string): ExcelShape;
  getItemOrNullObject?(key: string): ExcelShape;
  load(props: string): void;
}
