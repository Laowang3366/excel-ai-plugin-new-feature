import { useEffect } from "react";

export interface DocumentDismissOptions {
  onDismiss: () => void;
  ignoreSelectors?: string[];
}

export interface UseDocumentDismissOptions extends DocumentDismissOptions {
  active: boolean;
  pointerEvent?: "click" | "mousedown";
  closeOnEscape?: boolean;
}

export function createDocumentDismissHandlers(options: DocumentDismissOptions) {
  const { onDismiss, ignoreSelectors = [] } = options;

  return {
    handlePointerEvent(event: MouseEvent) {
      const target = event.target as Element | null;
      if (target && ignoreSelectors.some((selector) => target.closest(selector))) return;
      onDismiss();
    },
    handleKeyDown(event: KeyboardEvent) {
      if (event.key === "Escape") onDismiss();
    },
  };
}

export function useDocumentDismiss(options: UseDocumentDismissOptions): void {
  const {
    active,
    onDismiss,
    ignoreSelectors,
    pointerEvent = "click",
    closeOnEscape = true,
  } = options;

  useEffect(() => {
    if (!active) return;
    const handlers = createDocumentDismissHandlers({ onDismiss, ignoreSelectors });
    document.addEventListener(pointerEvent, handlers.handlePointerEvent);
    if (closeOnEscape) {
      document.addEventListener("keydown", handlers.handleKeyDown);
    }
    return () => {
      document.removeEventListener(pointerEvent, handlers.handlePointerEvent);
      if (closeOnEscape) {
        document.removeEventListener("keydown", handlers.handleKeyDown);
      }
    };
  }, [active, closeOnEscape, ignoreSelectors, onDismiss, pointerEvent]);
}
