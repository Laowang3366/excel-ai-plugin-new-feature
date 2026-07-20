import type {
  RangeAutofitInfo,
  RangeAutofitInput,
  RangeDeleteInput,
  RangeInsertInput,
  RangeMutationInfo,
} from "./rangeStructureTypes";
import type { HostResult } from "./types";
import { unsupported } from "./types";

const RANGE_STRUCTURE_EVIDENCE =
  "No in-repo WPS JSA Range.Insert/Delete or Range.AutoFit contract";

export async function wpsInsertRange(
  _input: RangeInsertInput,
): Promise<HostResult<RangeMutationInfo>> {
  return unsupported(
    "range.insert",
    "wps-jsa",
    "range.insert is not verified for WPS JSA",
    RANGE_STRUCTURE_EVIDENCE,
  );
}

export async function wpsDeleteRange(
  _input: RangeDeleteInput,
): Promise<HostResult<RangeMutationInfo>> {
  return unsupported(
    "range.delete",
    "wps-jsa",
    "range.delete is not verified for WPS JSA",
    RANGE_STRUCTURE_EVIDENCE,
  );
}

export async function wpsAutofitRange(
  _input: RangeAutofitInput,
): Promise<HostResult<RangeAutofitInfo>> {
  return unsupported(
    "range.autofit",
    "wps-jsa",
    "range.autofit is not verified for WPS JSA",
    RANGE_STRUCTURE_EVIDENCE,
  );
}
