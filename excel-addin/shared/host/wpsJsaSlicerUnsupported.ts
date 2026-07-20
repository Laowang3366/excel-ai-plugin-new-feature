/**
 * WPS JSA: no in-repo / official Slicer member evidence → typed unsupported for all tools.
 * COM/.NET/Shell fallback forbidden.
 */
import type {
  SlicerCreateInfo,
  SlicerCreateInput,
  SlicerDeleteInfo,
  SlicerDeleteInput,
  SlicerFilterApplyInput,
  SlicerFilterClearInput,
  SlicerFilterGetInput,
  SlicerFilterInfo,
  SlicerInfo,
  SlicerListInfo,
  SlicerListInput,
  SlicerUpdateInput,
} from "./slicerTypes";
import type { SlicerHostMethods } from "./slicerHostMethods";
import type { HostResult } from "./types";
import { unsupported } from "./types";

const EVIDENCE =
  "No verified WPS JSA Slicer/Slicers contract in this repository; COM/.NET/Shell fallback forbidden";

function unsup<T>(capability: string): Promise<HostResult<T>> {
  return Promise.resolve(
    unsupported(
      capability,
      "wps-jsa",
      "Slicers are not verified for WPS JSA in this repository",
      EVIDENCE,
    ) as HostResult<T>,
  );
}

export async function wpsListSlicers(_input?: SlicerListInput): Promise<HostResult<SlicerListInfo>> {
  return unsup("slicer.list");
}
export async function wpsCreateSlicer(_input: SlicerCreateInput): Promise<HostResult<SlicerCreateInfo>> {
  return unsup("slicer.create");
}
export async function wpsUpdateSlicer(_input: SlicerUpdateInput): Promise<HostResult<SlicerInfo>> {
  return unsup("slicer.update");
}
export async function wpsDeleteSlicer(_input: SlicerDeleteInput): Promise<HostResult<SlicerDeleteInfo>> {
  return unsup("slicer.delete");
}
export async function wpsGetSlicerFilter(
  _input: SlicerFilterGetInput,
): Promise<HostResult<SlicerFilterInfo>> {
  return unsup("slicer.filter.get");
}
export async function wpsApplySlicerFilter(
  _input: SlicerFilterApplyInput,
): Promise<HostResult<SlicerFilterInfo>> {
  return unsup("slicer.filter.apply");
}
export async function wpsClearSlicerFilter(
  _input: SlicerFilterClearInput,
): Promise<HostResult<SlicerFilterInfo>> {
  return unsup("slicer.filter.clear");
}

/**
 * Base class holding slicer method bindings so WpsJsaAdapter stays under the line budget.
 */
export abstract class WpsJsaSlicerSupport implements SlicerHostMethods {
  listSlicers = wpsListSlicers;
  createSlicer = wpsCreateSlicer;
  updateSlicer = wpsUpdateSlicer;
  deleteSlicer = wpsDeleteSlicer;
  getSlicerFilter = wpsGetSlicerFilter;
  applySlicerFilter = wpsApplySlicerFilter;
  clearSlicerFilter = wpsClearSlicerFilter;
}
