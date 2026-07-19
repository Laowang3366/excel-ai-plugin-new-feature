import type { HostResult } from "./types";
import { unsupported } from "./types";

const RANGE_IMAGE_EVIDENCE =
  "Range.getImage (ExcelApi 1.7) has no verified WPS JSA Base64 PNG export";

export async function wpsGetRangeImage(_input: unknown) {
  return unsupported(
    "range.image.get",
    "wps-jsa",
    "range.image.get is not verified for WPS JSA (no Base64 PNG export)",
    RANGE_IMAGE_EVIDENCE,
  ) as HostResult<never>;
}
