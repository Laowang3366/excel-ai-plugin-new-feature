/** Office.js Worksheet display props; empty tabColor = automatic. */
export interface SheetDisplayInfo {
  sheetName: string;
  tabColor: string;
  showGridlines: boolean;
  showHeadings: boolean;
}

export interface SheetDisplayUpdateInput {
  sheetName: string;
  tabColor?: string;
  showGridlines?: boolean;
  showHeadings?: boolean;
}
