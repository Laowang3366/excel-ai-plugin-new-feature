/**
 * Office.js slicer method group for OfficeJsAdapter binding.
 */
import {
  officeJsCreateSlicer,
  officeJsDeleteSlicer,
  officeJsListSlicers,
} from "./officeJsSlicer";
import {
  officeJsApplySlicerFilter,
  officeJsClearSlicerFilter,
  officeJsGetSlicerFilter,
} from "./officeJsSlicerFilter";
import { officeJsUpdateSlicer } from "./officeJsSlicerUpdate";
import type { SlicerHostMethods } from "./slicerHostMethods";

export const officeJsSlicerMethods: SlicerHostMethods = {
  listSlicers: officeJsListSlicers,
  createSlicer: officeJsCreateSlicer,
  updateSlicer: officeJsUpdateSlicer,
  deleteSlicer: officeJsDeleteSlicer,
  getSlicerFilter: officeJsGetSlicerFilter,
  applySlicerFilter: officeJsApplySlicerFilter,
  clearSlicerFilter: officeJsClearSlicerFilter,
};
