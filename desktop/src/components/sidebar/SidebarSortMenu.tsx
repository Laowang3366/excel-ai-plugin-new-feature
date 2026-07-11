import type { SidebarSortMode } from "../../utils/sidebarHelpers";
import {
  Check,
  Clock,
} from "../common/IconMap";

export type SidebarSortSection = "projects" | "conversations";

interface SidebarSortMenuLabels {
  recentDesc: string;
  recentAsc: string;
  nameAsc: string;
  nameDesc: string;
}

interface SidebarSortMenuProps {
  menu: { section: SidebarSortSection; x: number; y: number };
  labels: SidebarSortMenuLabels;
  projectSortMode: SidebarSortMode;
  conversationSortMode: SidebarSortMode;
  onSelectSortMode: (section: SidebarSortSection, mode: SidebarSortMode) => void;
}

const SORT_OPTIONS: Array<[SidebarSortMode, keyof SidebarSortMenuLabels]> = [
  ["recentDesc", "recentDesc"],
  ["recentAsc", "recentAsc"],
  ["nameAsc", "nameAsc"],
  ["nameDesc", "nameDesc"],
];

export function SidebarSortMenu({
  menu,
  labels,
  projectSortMode,
  conversationSortMode,
  onSelectSortMode,
}: SidebarSortMenuProps) {
  const activeMode = menu.section === "projects" ? projectSortMode : conversationSortMode;

  return (
    <div
      className="sidebar-sort-menu"
      style={{ left: menu.x, top: menu.y }}
      data-section={menu.section}
      onClick={(event) => event.stopPropagation()}
    >
      {SORT_OPTIONS.map(([mode, labelKey]) => (
        <button
          key={mode}
          className={`sidebar-sort-menu-item${activeMode === mode ? " active" : ""}`}
          onClick={() => onSelectSortMode(menu.section, mode)}
        >
          <Clock size={14} />
          <span>{labels[labelKey]}</span>
          {activeMode === mode && <Check size={14} />}
        </button>
      ))}
    </div>
  );
}
