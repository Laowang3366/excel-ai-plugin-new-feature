/**
 * 浮动任务面板 — 可拖拽、可折叠的通用浮窗容器
 *
 * 从 ChatPage.tsx 提取，零聊天业务依赖，
 * 仅依赖 getTaskPanelMeta 获取面板标题与图标。
 */

import React, { useRef, useEffect, useState, useCallback } from "react";
import { useSettingsStore } from "../../store/settingsStore";
import { getAppText } from "../../i18n";
import { getTaskPanelMeta, clamp, type ActiveIntentKind } from "../../utils/chatHelpers";
import { ChevronDown, Minimize2, X } from "../common/IconMap";

interface FloatingTaskPanelProps {
  intent: ActiveIntentKind;
  children: React.ReactNode;
  onClose: () => void;
}

export function FloatingTaskPanel({
  intent,
  children,
  onClose,
}: FloatingTaskPanelProps) {
  const { language } = useSettingsStore();
  const text = getAppText(language);
  const panelRef = useRef<HTMLDivElement>(null);
  const dragRef = useRef({ pointerId: -1, offsetX: 0, offsetY: 0 });
  const [collapsed, setCollapsed] = useState(false);
  const [position, setPosition] = useState({ x: 0, y: 64 });
  const meta = getTaskPanelMeta(intent, language);

  const getBounds = useCallback(() => {
    const panel = panelRef.current;
    const parent = panel?.parentElement;
    if (!panel || !parent) return null;
    const parentRect = parent.getBoundingClientRect();
    const panelRect = panel.getBoundingClientRect();
    return {
      width: parentRect.width,
      height: parentRect.height,
      panelWidth: panelRect.width || 520,
      panelHeight: panelRect.height || 64,
    };
  }, []);

  const constrainPosition = useCallback((next: { x: number; y: number }) => {
    const bounds = getBounds();
    if (!bounds) return next;
    return {
      x: clamp(next.x, 12, bounds.width - bounds.panelWidth - 12),
      y: clamp(next.y, 12, bounds.height - bounds.panelHeight - 12),
    };
  }, [getBounds]);

  useEffect(() => {
    setCollapsed(false);
    requestAnimationFrame(() => {
      const bounds = getBounds();
      if (!bounds) return;
      setPosition({
        x: clamp(bounds.width - bounds.panelWidth - 24, 12, bounds.width - bounds.panelWidth - 12),
        y: 62,
      });
    });
  }, [intent, getBounds]);

  useEffect(() => {
    const handleResize = () => setPosition((current) => constrainPosition(current));
    window.addEventListener("resize", handleResize);
    return () => window.removeEventListener("resize", handleResize);
  }, [constrainPosition]);

  useEffect(() => {
    requestAnimationFrame(() => {
      const bounds = getBounds();
      if (!bounds) return;
      setPosition((current) => {
        if (collapsed) {
          return {
            x: clamp(bounds.width - bounds.panelWidth - 24, 12, bounds.width - bounds.panelWidth - 12),
            y: clamp(current.y, 12, bounds.height - bounds.panelHeight - 12),
          };
        }
        return constrainPosition(current);
      });
    });
  }, [collapsed, constrainPosition, getBounds]);

  useEffect(() => {
    if (collapsed) return;

    const handleOutsidePointerDown = (event: PointerEvent) => {
      const panel = panelRef.current;
      const target = event.target;
      if (!panel || !(target instanceof Node) || panel.contains(target)) return;
      setCollapsed(true);
    };

    document.addEventListener("pointerdown", handleOutsidePointerDown, true);
    return () => document.removeEventListener("pointerdown", handleOutsidePointerDown, true);
  }, [collapsed]);

  const handlePointerDown = (event: React.PointerEvent<HTMLDivElement>) => {
    if (event.button !== 0) return;
    const panel = panelRef.current;
    const parent = panel?.parentElement;
    if (!panel || !parent) return;
    const parentRect = parent.getBoundingClientRect();
    dragRef.current = {
      pointerId: event.pointerId,
      offsetX: event.clientX - parentRect.left - position.x,
      offsetY: event.clientY - parentRect.top - position.y,
    };
    panel.setPointerCapture(event.pointerId);
  };

  const handlePointerMove = (event: React.PointerEvent<HTMLDivElement>) => {
    if (dragRef.current.pointerId < 0 || event.pointerId !== dragRef.current.pointerId || event.buttons !== 1) return;
    const parent = panelRef.current?.parentElement;
    if (!parent) return;
    const parentRect = parent.getBoundingClientRect();
    const next = {
      x: event.clientX - parentRect.left - dragRef.current.offsetX,
      y: event.clientY - parentRect.top - dragRef.current.offsetY,
    };
    setPosition(constrainPosition(next));
  };

  const handlePointerUp = (event: React.PointerEvent<HTMLDivElement>) => {
    dragRef.current.pointerId = -1;
    panelRef.current?.releasePointerCapture(event.pointerId);
  };

  return (
    <section
      ref={panelRef}
      className={`task-floating-panel ${collapsed ? "collapsed" : ""}`}
      style={{ transform: `translate(${position.x}px, ${position.y}px)` }}
      onPointerMove={handlePointerMove}
      onPointerUp={handlePointerUp}
      onPointerCancel={handlePointerUp}
    >
      <div
        className="task-floating-header"
        onPointerDown={handlePointerDown}
      >
        <div className="task-floating-title">
          {meta.icon}
          <span>{meta.title}</span>
        </div>
        <div className="task-floating-actions">
          <button
            className="task-floating-btn"
            onPointerDown={(event) => event.stopPropagation()}
            onClick={() => setCollapsed((value) => !value)}
            title={collapsed ? text.chat.expand : text.chat.collapse}
          >
            {collapsed ? <ChevronDown size={15} /> : <Minimize2 size={15} />}
          </button>
          <button
            className="task-floating-btn"
            onPointerDown={(event) => event.stopPropagation()}
            onClick={onClose}
            title={text.chat.close}
          >
            <X size={15} />
          </button>
        </div>
      </div>
      {!collapsed && (
        <div className="task-floating-body">
          {children}
        </div>
      )}
    </section>
  );
}
