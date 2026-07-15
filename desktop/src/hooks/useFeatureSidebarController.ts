import { useCallback, useReducer, useRef } from "react";

import type { IntentKind } from "../components/Sidebar";
import {
  INITIAL_FEATURE_SIDEBAR_STATE,
  reduceFeatureSidebarState,
  shouldFocusFeatureSidebarOnToggle,
  shouldRestoreFeatureSidebarFocus,
  type FeatureSidebarCloseReason,
} from "../utils/featureSidebarState";

export function useFeatureSidebarController() {
  const [state, dispatch] = useReducer(reduceFeatureSidebarState, INITIAL_FEATURE_SIDEBAR_STATE);
  const toggleRef = useRef<HTMLButtonElement>(null);

  const close = useCallback(
    (reason: FeatureSidebarCloseReason) => {
      if (!state.isOpen) return;
      dispatch({ type: "close" });
      if (shouldRestoreFeatureSidebarFocus(reason)) {
        window.requestAnimationFrame(() => toggleRef.current?.focus());
      }
    },
    [state.isOpen],
  );

  const toggle = useCallback(() => {
    const focusFirstShortcut = shouldFocusFeatureSidebarOnToggle(state.isOpen);
    dispatch({ type: "toggle" });
    if (focusFirstShortcut) {
      window.requestAnimationFrame(() => {
        toggleRef.current
          ?.closest(".chat-page")
          ?.querySelector<HTMLButtonElement>(".feature-sidebar-shortcut")
          ?.focus();
      });
    }
  }, [state.isOpen]);

  const select = useCallback((intent: NonNullable<IntentKind>) => {
    dispatch({ type: "select", intent });
  }, []);

  const closeAfterSend = useCallback(() => {
    close("send");
  }, [close]);

  const closeManually = useCallback(() => {
    close("manual");
  }, [close]);

  return {
    activeIntent: state.activeIntent,
    closeAfterSend,
    closeManually,
    isOpen: state.isOpen,
    select,
    toggle,
    toggleRef,
  };
}
