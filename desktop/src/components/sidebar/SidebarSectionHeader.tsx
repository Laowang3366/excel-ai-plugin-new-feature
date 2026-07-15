import React from "react";
import { ChevronDown, ChevronRight, ClipboardList } from "../common/IconMap";

interface SidebarSectionHeaderProps {
  title: string;
  expanded: boolean;
  sortTitle: string;
  actionTitle: string;
  actionIcon: React.ReactNode;
  actionClassName?: string;
  actionDisabled?: boolean;
  onToggle: () => void;
  onOpenSort: (event: React.MouseEvent) => void;
  onAction: () => void;
}

export function SidebarSectionHeader({
  title,
  expanded,
  sortTitle,
  actionTitle,
  actionIcon,
  actionClassName = "",
  actionDisabled = false,
  onToggle,
  onOpenSort,
  onAction,
}: SidebarSectionHeaderProps) {
  return (
    <div className="sidebar-section-header">
      <button className="sidebar-section-toggle" onClick={onToggle}>
        <span>{title}</span>
        {expanded ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
      </button>
      <button className="sidebar-section-add" onClick={onOpenSort} title={sortTitle}>
        <ClipboardList size={14} />
      </button>
      <button
        className={`sidebar-section-add${actionClassName}`}
        onClick={onAction}
        disabled={actionDisabled}
        title={actionTitle}
      >
        {actionIcon}
      </button>
    </div>
  );
}
