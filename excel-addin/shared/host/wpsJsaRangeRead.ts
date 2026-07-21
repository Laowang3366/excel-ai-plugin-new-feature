import {
  formulaMatrixFrom,
  getSheet,
  matrixFrom,
  requireWorkbook,
  type WpsRange,
} from "./wpsJsaRuntime";
import type { HostResult, RangeData, RangeExpandMode } from "./types";
import { ok, unsupported } from "./types";
import { hasWpsAddressSurface, readWpsAddress } from "./wpsJsaAddress";

const RANGE_EVIDENCE =
  "Assumed Worksheets.Item(name).Range(address).Value2 (not in bridge contract)";
const CURRENT_REGION_EVIDENCE =
  "Assumed Range.CurrentRegion (ET COM parity; not in bridge contract; not device-verified)";

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

/**
 * range.read for WPS JSA.
 * expand: none | currentRegion (member-probed). spill/currentArray → typed unsupported.
 * usedRegion is not part of RangeExpandMode; sheet UsedRange stays on formula.context/inspect.
 */
export async function wpsReadRange(
  sheetName: string,
  address: string,
  expand?: RangeExpandMode,
): Promise<HostResult<RangeData>> {
  const bare = address.includes("!") ? address.split("!")[1]! : address;
  const isSingle = !bare.includes(":") && !bare.includes(",");
  // Desktop/public contract: omitted expand on single cell means spill.
  const effectiveExpand =
    expand === undefined && isSingle ? ("spill" as const) : expand;
  if (effectiveExpand === "spill" || effectiveExpand === "currentArray") {
    return unsupported(
      "range.read",
      "wps-jsa",
      `expand "${effectiveExpand}" is not verified for WPS JSA`,
      "No in-repo spill/currentArray contract",
    );
  }

  const workbookResult = requireWorkbook("range.read");
  if (!workbookResult.ok) return workbookResult;
  const sheet = getSheet(workbookResult.data, sheetName);
  if (!sheet?.Range) {
    return unsupported(
      "range.read",
      "wps-jsa",
      `Sheet "${sheetName}" or Range API missing`,
      RANGE_EVIDENCE,
    );
  }

  let range: WpsRange;
  try {
    range = sheet.Range(address);
  } catch (error) {
    return unsupported("range.read", "wps-jsa", messageOf(error), RANGE_EVIDENCE);
  }

  let expanded = false;
  let expandMode: RangeExpandMode = "none";
  if (effectiveExpand === "currentRegion") {
    let region: WpsRange | undefined;
    try {
      region = range.CurrentRegion;
    } catch (error) {
      return unsupported(
        "range.read",
        "wps-jsa",
        `Range.CurrentRegion access failed: ${messageOf(error)}`,
        CURRENT_REGION_EVIDENCE,
      );
    }
    if (!region || !hasWpsAddressSurface(region)) {
      return unsupported(
        "range.read",
        "wps-jsa",
        "Range.CurrentRegion is unavailable",
        CURRENT_REGION_EVIDENCE,
      );
    }
    range = region;
    expanded = true;
    expandMode = "currentRegion";
  }

  try {
    return ok({
      sheetName,
      address: readWpsAddress(range, address) ?? address,
      values: matrixFrom(range.Value2),
      formulas: formulaMatrixFrom(range.Formula),
      expanded,
      expandMode,
    });
  } catch (error) {
    return unsupported(
      "range.read",
      "wps-jsa",
      messageOf(error),
      "Assumed Range.Value2/Formula (not in bridge contract)",
    );
  }
}
