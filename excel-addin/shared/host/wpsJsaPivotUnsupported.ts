/**
 * WPS JSA: no in-repo / official PivotTables contract → typed unsupported.
 */
import type {
  PivotCreateInfo,
  PivotCreateInput,
  PivotListInfo,
  PivotListInput,
  PivotRefreshInfo,
  PivotRefreshInput,
} from "./pivotTypes";
import type { HostResult } from "./types";
import { unsupported } from "./types";

const EVIDENCE =
  "No verified WPS JSA PivotTables/PivotCaches contract in this repository; COM/.NET/Shell fallback forbidden";

export async function wpsListPivots(_input?: PivotListInput): Promise<HostResult<PivotListInfo>> {
  return unsupported(
    "pivot.list",
    "wps-jsa",
    "PivotTables are not verified for WPS JSA in this repository",
    EVIDENCE,
  );
}

export async function wpsCreatePivot(_input: PivotCreateInput): Promise<HostResult<PivotCreateInfo>> {
  return unsupported(
    "pivot.create",
    "wps-jsa",
    "PivotTables are not verified for WPS JSA in this repository",
    EVIDENCE,
  );
}

export async function wpsRefreshPivots(
  _input?: PivotRefreshInput,
): Promise<HostResult<PivotRefreshInfo>> {
  return unsupported(
    "pivot.refresh",
    "wps-jsa",
    "PivotTables are not verified for WPS JSA in this repository",
    EVIDENCE,
  );
}
