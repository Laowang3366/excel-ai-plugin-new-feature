/**
 * Slicer domain method group for HostAdapter (keeps hostAdapter.ts under line budget).
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
import type { HostResult } from "./types";

export interface SlicerHostMethods {
  listSlicers(input?: SlicerListInput): Promise<HostResult<SlicerListInfo>>;
  createSlicer(input: SlicerCreateInput): Promise<HostResult<SlicerCreateInfo>>;
  updateSlicer(input: SlicerUpdateInput): Promise<HostResult<SlicerInfo>>;
  deleteSlicer(input: SlicerDeleteInput): Promise<HostResult<SlicerDeleteInfo>>;
  getSlicerFilter(input: SlicerFilterGetInput): Promise<HostResult<SlicerFilterInfo>>;
  applySlicerFilter(input: SlicerFilterApplyInput): Promise<HostResult<SlicerFilterInfo>>;
  clearSlicerFilter(input: SlicerFilterClearInput): Promise<HostResult<SlicerFilterInfo>>;
}
