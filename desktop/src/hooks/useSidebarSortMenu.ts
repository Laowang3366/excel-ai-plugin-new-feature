import { useCallback, useState } from "react";
import type { MouseEvent } from "react";
import type { SidebarSortMode, SidebarSortSection } from "../components/sidebar/SidebarExpanded";

export function useSidebarSortMenu(onOpenSortMenu: () => void) {
  const [sortMenu, setSortMenu] = useState<{
    section: SidebarSortSection;
    x: number;
    y: number;
  } | null>(null);
  const [projectSortMode, setProjectSortMode] = useState<SidebarSortMode>("recentDesc");
  const [conversationSortMode, setConversationSortMode] = useState<SidebarSortMode>("recentDesc");

  const handleOpenSortMenu = useCallback(
    (event: MouseEvent, section: SidebarSortSection) => {
      event.stopPropagation();
      onOpenSortMenu();
      const rect = event.currentTarget.getBoundingClientRect();
      const menuWidth = 168;
      setSortMenu({
        section,
        x: Math.max(8, Math.min(window.innerWidth - menuWidth - 8, rect.right - menuWidth)),
        y: Math.max(8, Math.min(window.innerHeight - 164, rect.bottom + 6)),
      });
    },
    [onOpenSortMenu],
  );

  const handleSelectSortMode = useCallback((section: SidebarSortSection, mode: SidebarSortMode) => {
    if (section === "projects") {
      setProjectSortMode(mode);
    } else {
      setConversationSortMode(mode);
    }
    setSortMenu(null);
  }, []);

  const closeSortMenu = useCallback(() => {
    setSortMenu(null);
  }, []);

  return {
    sortMenu,
    projectSortMode,
    conversationSortMode,
    handleOpenSortMenu,
    handleSelectSortMode,
    closeSortMenu,
  };
}
