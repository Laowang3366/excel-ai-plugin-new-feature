import { useCallback, useRef, useState } from "react";
import type { MouseEvent as ReactMouseEvent } from "react";

export function useSidebarResize(initialWidth = 260) {
  const [sidebarWidth, setSidebarWidth] = useState(initialWidth);
  const [isResizing, setIsResizing] = useState(false);
  const resizingRef = useRef(false);

  const handleResizeStart = useCallback(
    (event: ReactMouseEvent) => {
      event.preventDefault();
      resizingRef.current = true;
      setIsResizing(true);
      const startX = event.clientX;
      const startWidth = sidebarWidth;
      let frameId: number | null = null;
      let nextWidth = startWidth;
      const flushWidth = () => {
        frameId = null;
        setSidebarWidth(nextWidth);
      };
      const handleMouseMove = (moveEvent: MouseEvent) => {
        if (!resizingRef.current) return;
        nextWidth = Math.min(400, Math.max(180, startWidth + moveEvent.clientX - startX));
        if (frameId === null) frameId = window.requestAnimationFrame(flushWidth);
      };
      const handleMouseUp = () => {
        resizingRef.current = false;
        setIsResizing(false);
        if (frameId !== null) {
          window.cancelAnimationFrame(frameId);
          frameId = null;
          setSidebarWidth(nextWidth);
        }
        document.removeEventListener("mousemove", handleMouseMove);
        document.removeEventListener("mouseup", handleMouseUp);
      };
      document.addEventListener("mousemove", handleMouseMove);
      document.addEventListener("mouseup", handleMouseUp);
    },
    [sidebarWidth],
  );

  return {
    sidebarWidth,
    isResizing,
    handleResizeStart,
  };
}
