import React, { useCallback, useEffect, useRef, useState } from "react";
import { useSettingsStore } from "../../store/settingsStore";
import { getAppText } from "../../i18n";
import { INTENT_SHORTCUTS, type IntentKind } from "../../utils/sidebarHelpers";
import { ChevronDown, Sparkles } from "./IconMap";
import {
  FEATURE_DOCK_COLLAPSED_SIZE,
  constrainFeatureDockAnchorPosition,
  getFeatureDockInitialPosition,
  getFeatureDockPointerAction,
  getFeatureDockResizePosition,
  shouldCollapseFeatureDockOnPointerDown,
} from "./featureFloatingDockGeometry";

export {
  constrainFeatureDockAnchorPosition,
  getFeatureDockInitialPosition,
  getFeatureDockPointerAction,
  getFeatureDockResizePosition,
  shouldCollapseFeatureDockOnPointerDown,
} from "./featureFloatingDockGeometry";

interface FeatureFloatingDockProps {
  activeIntent: IntentKind;
  onIntentClick: (intent: IntentKind) => void;
}

export function FeatureFloatingDock({
  activeIntent,
  onIntentClick,
}: FeatureFloatingDockProps) {
  const { language } = useSettingsStore();
  const text = getAppText(language);
  const dockRef = useRef<HTMLDivElement>(null);
  const dragRef = useRef({
    pointerId: -1,
    offsetX: 0,
    offsetY: 0,
    startX: 0,
    startY: 0,
    expandedAtPointerDown: false,
    moved: false,
  });
  const suppressClickRef = useRef(false);
  const userMovedRef = useRef(false);
  const [expanded, setExpanded] = useState(false);
  const [position, setPosition] = useState({ x: 0, y: 140 });

  const getBounds = useCallback((nextExpanded: boolean) => {
    const dock = dockRef.current;
    const parent = dock?.parentElement;
    if (!dock || !parent) return null;
    const parentRect = parent.getBoundingClientRect();
    const dockRect = dock.getBoundingClientRect();
    return {
      width: parentRect.width,
      height: parentRect.height,
      dockHeight: dockRect.height || (nextExpanded ? 288 : FEATURE_DOCK_COLLAPSED_SIZE),
    };
  }, []);

  const constrainPosition = useCallback((next: { x: number; y: number }, nextExpanded = expanded) => {
    const bounds = getBounds(nextExpanded);
    if (!bounds) return next;
    return constrainFeatureDockAnchorPosition(next, bounds.width, bounds.height, nextExpanded);
  }, [expanded, getBounds]);

  useEffect(() => {
    requestAnimationFrame(() => {
      const bounds = getBounds(false);
      if (!bounds) return;
      setPosition(getFeatureDockInitialPosition(bounds));
    });
  }, [getBounds]);

  useEffect(() => {
    const updatePositionForBounds = () => {
      const bounds = getBounds(false);
      if (!bounds) {
        setPosition((current) => constrainPosition(current));
        return;
      }
      setPosition((current) => getFeatureDockResizePosition({
        current,
        width: bounds.width,
        height: bounds.height,
        userMoved: userMovedRef.current,
        expanded,
      }));
    };
    const parent = dockRef.current?.parentElement;
    let resizeObserver: ResizeObserver | null = null;
    if (parent && typeof ResizeObserver !== "undefined") {
      resizeObserver = new ResizeObserver(updatePositionForBounds);
      resizeObserver.observe(parent);
    }
    window.addEventListener("resize", updatePositionForBounds);
    return () => {
      resizeObserver?.disconnect();
      window.removeEventListener("resize", updatePositionForBounds);
    };
  }, [constrainPosition, expanded, getBounds]);

  const toggleExpanded = useCallback(() => {
    const nextExpanded = !expanded;
    setPosition((current) => constrainPosition(current, nextExpanded));
    setExpanded(nextExpanded);
  }, [expanded, constrainPosition]);

  const collapseToBall = useCallback(() => {
    setPosition((current) => constrainPosition(current, false));
    setExpanded(false);
  }, [constrainPosition]);

  useEffect(() => {
    if (!expanded) return;

    const handleOutsidePointerDown = (event: PointerEvent) => {
      const dock = dockRef.current;
      const target = event.target;
      const targetInsideDock = !!dock && target instanceof Node && dock.contains(target);
      if (shouldCollapseFeatureDockOnPointerDown({ expanded: true, targetInsideDock })) {
        collapseToBall();
      }
    };

    document.addEventListener("pointerdown", handleOutsidePointerDown, true);
    return () => document.removeEventListener("pointerdown", handleOutsidePointerDown, true);
  }, [collapseToBall, expanded]);

  const handlePointerDown = (event: React.PointerEvent<HTMLElement>) => {
    if (event.button !== 0) return;
    const dock = dockRef.current;
    const parent = dock?.parentElement;
    if (!dock || !parent) return;
    const parentRect = parent.getBoundingClientRect();
    dragRef.current = {
      pointerId: event.pointerId,
      offsetX: event.clientX - parentRect.left - position.x,
      offsetY: event.clientY - parentRect.top - position.y,
      startX: event.clientX,
      startY: event.clientY,
      expandedAtPointerDown: expanded,
      moved: false,
    };
    dock.setPointerCapture(event.pointerId);
  };

  const handlePointerMove = (event: React.PointerEvent<HTMLDivElement>) => {
    const drag = dragRef.current;
    if (drag.pointerId < 0 || event.pointerId !== drag.pointerId || event.buttons !== 1) return;
    if (Math.abs(event.clientX - drag.startX) > 3 || Math.abs(event.clientY - drag.startY) > 3) {
      drag.moved = true;
    }
    const parent = dockRef.current?.parentElement;
    if (!parent) return;
    const parentRect = parent.getBoundingClientRect();
    setPosition(constrainPosition({
      x: event.clientX - parentRect.left - drag.offsetX,
      y: event.clientY - parentRect.top - drag.offsetY,
    }));
  };

  const handlePointerUp = (event: React.PointerEvent<HTMLDivElement>) => {
    const action = getFeatureDockPointerAction({
      activePointerId: dragRef.current.pointerId,
      eventPointerId: event.pointerId,
      expandedAtPointerDown: dragRef.current.expandedAtPointerDown,
      moved: dragRef.current.moved,
    });
    dragRef.current.pointerId = -1;
    if (dockRef.current?.hasPointerCapture(event.pointerId)) {
      dockRef.current.releasePointerCapture(event.pointerId);
    }
    if (action === "toggle") {
      toggleExpanded();
      return;
    }
    if (action === "drag") {
      userMovedRef.current = true;
      suppressClickRef.current = true;
      window.setTimeout(() => {
        suppressClickRef.current = false;
      }, 0);
    }
  };

  const handlePointerCancel = (event: React.PointerEvent<HTMLDivElement>) => {
    dragRef.current.pointerId = -1;
    if (dockRef.current?.hasPointerCapture(event.pointerId)) {
      dockRef.current.releasePointerCapture(event.pointerId);
    }
  };

  const handleToggleClick = () => {
    if (suppressClickRef.current) return;
    toggleExpanded();
  };

  return (
    <div
      ref={dockRef}
      className={`feature-floating-dock ${expanded ? "expanded" : ""}`}
      style={{ transform: `translate(${position.x}px, ${position.y}px)` }}
      onPointerMove={handlePointerMove}
      onPointerUp={handlePointerUp}
      onPointerCancel={handlePointerCancel}
    >
      {expanded ? (
        <section className="feature-floating-card">
          <div className="feature-floating-header" onPointerDown={handlePointerDown}>
            <div className="feature-floating-title">
              <Sparkles size={16} />
              <span>{text.sidebar.features}</span>
            </div>
            <button
              className="feature-floating-collapse"
              type="button"
              onPointerDown={(event) => event.stopPropagation()}
              onClick={handleToggleClick}
              title={text.chat.collapse}
            >
              <ChevronDown size={15} />
            </button>
          </div>
          <div className="feature-floating-list">
            {INTENT_SHORTCUTS.map((shortcut) => (
              <button
                key={shortcut.key}
                className={`feature-floating-item ${activeIntent === shortcut.key ? "active" : ""}`}
                type="button"
                onClick={() => {
                  onIntentClick(shortcut.key);
                  collapseToBall();
                }}
              >
                <shortcut.icon size={16} />
                <span>{text.sidebar.intents[shortcut.key]}</span>
              </button>
            ))}
          </div>
        </section>
      ) : (
        <button
          className="feature-floating-ball"
          type="button"
          onPointerDown={handlePointerDown}
          title={text.sidebar.features}
        >
          <Sparkles size={20} />
        </button>
      )}
    </div>
  );
}
