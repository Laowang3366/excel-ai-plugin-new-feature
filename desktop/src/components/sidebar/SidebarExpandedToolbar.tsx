import React from "react";
import type { getAppText } from "../../i18n";
import { PenLine, RefreshCw, Search } from "../common/IconMap";

interface SidebarExpandedToolbarProps {
  text: ReturnType<typeof getAppText>;
  creatingNewThread: boolean;
  searchOpen: boolean;
  onCreateNewThread: () => void;
  onToggleSearch: () => void;
}

export function SidebarExpandedToolbar({
  text,
  creatingNewThread,
  searchOpen,
  onCreateNewThread,
  onToggleSearch,
}: SidebarExpandedToolbarProps) {
  return (
    <div className="sidebar-primary-nav">
      <button
        className={`sidebar-primary-action${creatingNewThread ? " creating" : ""}`}
        onClick={onCreateNewThread}
        disabled={creatingNewThread}
        title={text.sidebar.newThread}
      >
        {creatingNewThread ? <RefreshCw size={16} className="spin" /> : <PenLine size={16} />}
        <span>{text.sidebar.newThread}</span>
      </button>
      <button
        className={`sidebar-primary-action${searchOpen ? " active" : ""}`}
        onClick={onToggleSearch}
        title={text.sidebar.search}
      >
        <Search size={16} />
        <span>{text.sidebar.search}</span>
      </button>
    </div>
  );
}
