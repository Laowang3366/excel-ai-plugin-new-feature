import { useEffect } from "react";

export interface DocumentDismissBoundaryRef {
  current: Element | null;
}

export interface DocumentDismissOptions {
  onDismiss: () => void;
  ignoreSelectors?: string[];
  boundaryRefs?: DocumentDismissBoundaryRef[];
}

export interface UseDocumentDismissOptions extends DocumentDismissOptions {
  active: boolean;
  pointerEvent?: "click" | "mousedown";
  closeOnEscape?: boolean;
}

export function createDocumentDismissHandlers(options: DocumentDismissOptions) {
  const { onDismiss, ignoreSelectors = [], boundaryRefs = [] } = options;

  return {
    handlePointerEvent(event: MouseEvent) {
      const target = event.target as EventTarget | null;
      const closest = (target as Element | null)?.closest;
      if (target && typeof closest === "function" && ignoreSelectors.some((selector) => closest.call(target, selector))) return;
      if (target && boundaryRefs.some((ref) => ref.current?.contains(target as Node))) return;
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
    boundaryRefs,
    ignoreSelectors,
    pointerEvent = "click",
    closeOnEscape = true,
  } = options;

  useEffect(() => {
    if (!active) return;
    const handlers = createDocumentDismissHandlers({ onDismiss, ignoreSelectors, boundaryRefs });
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
  }, [active, boundaryRefs, closeOnEscape, ignoreSelectors, onDismiss, pointerEvent]);
}
