import { describe, expect, test } from "vitest";
import {
  constrainFeatureDockAnchorPosition,
  getFeatureDockInitialPosition,
  getFeatureDockPointerAction,
  getFeatureDockResizePosition,
  shouldCollapseFeatureDockOnPointerDown,
} from "./FeatureFloatingDock";

describe("getFeatureDockPointerAction", () => {
  test("toggles the dock when the collapsed ball is released without dragging", () => {
    expect(getFeatureDockPointerAction({ expandedAtPointerDown: false, moved: false })).toBe("toggle");
  });

  test("does not toggle the dock after dragging the collapsed ball", () => {
    expect(getFeatureDockPointerAction({ expandedAtPointerDown: false, moved: true })).toBe("drag");
  });

  test("does not toggle the expanded header when it is released without dragging", () => {
    expect(getFeatureDockPointerAction({ expandedAtPointerDown: true, moved: false })).toBe("none");
  });

  test("ignores pointerup events that do not belong to the dock drag gesture", () => {
    expect(
      getFeatureDockPointerAction({
        activePointerId: -1,
        eventPointerId: 7,
        expandedAtPointerDown: false,
        moved: false,
      })
    ).toBe("none");
  });
});

describe("shouldCollapseFeatureDockOnPointerDown", () => {
  test("collapses when the expanded dock receives an outside pointerdown", () => {
    expect(shouldCollapseFeatureDockOnPointerDown({ expanded: true, targetInsideDock: false })).toBe(true);
  });

  test("keeps the dock open when clicking inside it", () => {
    expect(shouldCollapseFeatureDockOnPointerDown({ expanded: true, targetInsideDock: true })).toBe(false);
  });

  test("does not handle outside pointerdown while already collapsed", () => {
    expect(shouldCollapseFeatureDockOnPointerDown({ expanded: false, targetInsideDock: false })).toBe(false);
  });
});

describe("getFeatureDockInitialPosition", () => {
  test("places the collapsed dock near the upper-right edge by default", () => {
    expect(getFeatureDockInitialPosition({ width: 1200, height: 900 })).toEqual({
      x: 1144,
      y: 70,
    });
  });
});

describe("getFeatureDockResizePosition", () => {
  test("keeps the default dock anchored to the upper-right after resizing wider", () => {
    expect(
      getFeatureDockResizePosition({
        current: { x: 1144, y: 70 },
        width: 2000,
        height: 1000,
        userMoved: false,
      })
    ).toEqual({ x: 1944, y: 70 });
  });

  test("keeps the default dock visible when the parent narrows", () => {
    expect(
      getFeatureDockResizePosition({
        current: { x: 1144, y: 70 },
        width: 900,
        height: 900,
        userMoved: false,
      })
    ).toEqual({ x: 844, y: 70 });
  });

  test("keeps a manually moved dock constrained instead of resetting it", () => {
    expect(
      getFeatureDockResizePosition({
        current: { x: 320, y: 180 },
        width: 2000,
        height: 1000,
        userMoved: true,
      })
    ).toEqual({ x: 320, y: 180 });
  });

  test("keeps an expanded manually moved dock inside the left and right edges", () => {
    expect(
      getFeatureDockResizePosition({
        current: { x: 40, y: 180 },
        width: 900,
        height: 900,
        userMoved: true,
        expanded: true,
      })
    ).toEqual({ x: 156, y: 180 });
  });
});

describe("constrainFeatureDockAnchorPosition", () => {
  test("anchors the expanded card to the ball without letting the card overflow right", () => {
    expect(
      constrainFeatureDockAnchorPosition(
        { x: 1144, y: 70 },
        1200,
        900,
        true,
      )
    ).toEqual({ x: 1144, y: 70 });
  });

  test("keeps the expanded card left edge visible", () => {
    expect(
      constrainFeatureDockAnchorPosition(
        { x: 24, y: 70 },
        1200,
        900,
        true,
      )
    ).toEqual({ x: 156, y: 70 });
  });
});
