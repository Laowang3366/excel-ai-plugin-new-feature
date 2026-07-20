import { useCallback, useEffect, useRef, type RefObject, type UIEvent } from "react";

const BOTTOM_THRESHOLD_PX = 80;

/**
 * Auto-scroll a container to bottom when `deps` change, unless the user has
 * scrolled up to read history (distance from bottom > threshold).
 */
export function useStickToBottom(deps: readonly unknown[]): {
  containerRef: RefObject<HTMLDivElement | null>;
  onScroll: (event: UIEvent<HTMLDivElement>) => void;
  pinToBottom: () => void;
} {
  const containerRef = useRef<HTMLDivElement | null>(null);
  const stickRef = useRef(true);

  const onScroll = useCallback((event: UIEvent<HTMLDivElement>) => {
    const el = event.currentTarget;
    const distance = el.scrollHeight - el.scrollTop - el.clientHeight;
    stickRef.current = distance <= BOTTOM_THRESHOLD_PX;
  }, []);

  const pinToBottom = useCallback(() => {
    const el = containerRef.current;
    if (!el || !stickRef.current) return;
    el.scrollTop = el.scrollHeight;
  }, []);

  useEffect(() => {
    pinToBottom();
    // Caller passes the reactive inputs that should re-pin when sticking.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, deps);

  return { containerRef, onScroll, pinToBottom };
}
