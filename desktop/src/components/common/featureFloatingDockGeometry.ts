import { clamp } from "../../utils/chatHelpers";

export const FEATURE_DOCK_COLLAPSED_SIZE = 44;
export const FEATURE_DOCK_EXPANDED_WIDTH = 188;
export const FEATURE_DOCK_EDGE_GAP = 12;
export const FEATURE_DOCK_DEFAULT_TOP = 70;

export type FeatureDockPointerAction = "toggle" | "drag" | "none";

export function getFeatureDockPointerAction({
  activePointerId,
  eventPointerId,
  expandedAtPointerDown,
  moved,
}: {
  activePointerId?: number;
  eventPointerId?: number;
  expandedAtPointerDown: boolean;
  moved: boolean;
}): FeatureDockPointerAction {
  if (
    activePointerId !== undefined &&
    eventPointerId !== undefined &&
    (activePointerId < 0 || activePointerId !== eventPointerId)
  ) {
    return "none";
  }
  if (moved) return "drag";
  return expandedAtPointerDown ? "none" : "toggle";
}

export function shouldCollapseFeatureDockOnPointerDown({
  expanded,
  targetInsideDock,
}: {
  expanded: boolean;
  targetInsideDock: boolean;
}): boolean {
  return expanded && !targetInsideDock;
}

export function getFeatureDockInitialPosition({
  width,
  height,
}: {
  width: number;
  height: number;
}) {
  return {
    x: clamp(
      width - FEATURE_DOCK_COLLAPSED_SIZE - FEATURE_DOCK_EDGE_GAP,
      12,
      width - FEATURE_DOCK_COLLAPSED_SIZE - 12,
    ),
    y: clamp(FEATURE_DOCK_DEFAULT_TOP, 54, height - FEATURE_DOCK_COLLAPSED_SIZE - 16),
  };
}

export function getFeatureDockResizePosition({
  current,
  width,
  height,
  userMoved,
  expanded = false,
}: {
  current: { x: number; y: number };
  width: number;
  height: number;
  userMoved: boolean;
  expanded?: boolean;
}) {
  if (!userMoved) {
    return getFeatureDockInitialPosition({ width, height });
  }
  return constrainFeatureDockAnchorPosition(current, width, height, expanded);
}

export function constrainFeatureDockAnchorPosition(
  current: { x: number; y: number },
  width: number,
  height: number,
  expanded: boolean,
) {
  const minX = expanded
    ? FEATURE_DOCK_EDGE_GAP + FEATURE_DOCK_EXPANDED_WIDTH - FEATURE_DOCK_COLLAPSED_SIZE
    : FEATURE_DOCK_EDGE_GAP;
  return {
    x: clamp(current.x, minX, width - FEATURE_DOCK_COLLAPSED_SIZE - FEATURE_DOCK_EDGE_GAP),
    y: clamp(current.y, 54, height - FEATURE_DOCK_COLLAPSED_SIZE - 16),
  };
}
