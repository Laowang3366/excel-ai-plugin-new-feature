import { describe, expect, it } from "vitest";
import {
  INITIAL_FEATURE_SIDEBAR_STATE,
  reduceFeatureSidebarState,
  shouldFocusFeatureSidebarOnToggle,
  shouldRestoreFeatureSidebarFocus,
} from "./featureSidebarState";

describe("reduceFeatureSidebarState", () => {
  it("opens without an active intent when toggled from the initial state", () => {
    expect(reduceFeatureSidebarState(INITIAL_FEATURE_SIDEBAR_STATE, { type: "toggle" })).toEqual({
      isOpen: true,
      activeIntent: null,
    });
  });

  it("opens with formula active when formula is selected", () => {
    expect(
      reduceFeatureSidebarState(INITIAL_FEATURE_SIDEBAR_STATE, {
        type: "select",
        intent: "formula",
      }),
    ).toEqual({
      isOpen: true,
      activeIntent: "formula",
    });
  });

  it("returns to the initial state when chart is closed", () => {
    const openChartState = reduceFeatureSidebarState(INITIAL_FEATURE_SIDEBAR_STATE, {
      type: "select",
      intent: "chart",
    });

    expect(reduceFeatureSidebarState(openChartState, { type: "close" })).toEqual(
      INITIAL_FEATURE_SIDEBAR_STATE,
    );
  });

  it("returns to the initial state when clean is toggled closed", () => {
    const openCleanState = reduceFeatureSidebarState(INITIAL_FEATURE_SIDEBAR_STATE, {
      type: "select",
      intent: "clean",
    });

    expect(reduceFeatureSidebarState(openCleanState, { type: "toggle" })).toEqual(
      INITIAL_FEATURE_SIDEBAR_STATE,
    );
  });

  it("restores toggle focus only after a manual close", () => {
    expect(shouldRestoreFeatureSidebarFocus("manual")).toBe(true);
    expect(shouldRestoreFeatureSidebarFocus("send")).toBe(false);
  });

  it("moves focus into the sidebar only when the toggle opens it", () => {
    expect(shouldFocusFeatureSidebarOnToggle(false)).toBe(true);
    expect(shouldFocusFeatureSidebarOnToggle(true)).toBe(false);
  });
});
