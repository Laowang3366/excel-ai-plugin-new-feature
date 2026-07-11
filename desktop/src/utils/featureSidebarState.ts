import type { IntentKind } from "./sidebarHelpers";

export interface FeatureSidebarState {
  isOpen: boolean;
  activeIntent: IntentKind;
}

export type FeatureSidebarAction =
  { type: "toggle" } | { type: "select"; intent: NonNullable<IntentKind> } | { type: "close" };

export const INITIAL_FEATURE_SIDEBAR_STATE: FeatureSidebarState = {
  isOpen: false,
  activeIntent: null,
};

export type FeatureSidebarCloseReason = "manual" | "send";

export function shouldRestoreFeatureSidebarFocus(reason: FeatureSidebarCloseReason): boolean {
  return reason === "manual";
}

export function shouldFocusFeatureSidebarOnToggle(isOpen: boolean): boolean {
  return !isOpen;
}

export function reduceFeatureSidebarState(
  state: FeatureSidebarState,
  action: FeatureSidebarAction,
): FeatureSidebarState {
  if (action.type === "select") {
    return { isOpen: true, activeIntent: action.intent };
  }
  if (action.type === "close" || state.isOpen) {
    return INITIAL_FEATURE_SIDEBAR_STATE;
  }
  return { isOpen: true, activeIntent: null };
}
