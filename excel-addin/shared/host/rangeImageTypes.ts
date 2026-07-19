/** Input for range.image.get (ExcelApi 1.7 Range.getImage). */
export interface RangeImageGetInput {
  sheetName: string;
  range: string;
}

/**
 * Host-generated range PNG snapshot (memory Base64 only).
 * No path, PDF, or MIME claim.
 */
export interface RangeImageInfo {
  /** Worksheet.name after load (host truth, not input echo). */
  sheetName: string;
  /** Range.address after load (host address text). */
  address: string;
  /** ClientResult.value from Range.getImage after sync. */
  imageBase64: string;
}
