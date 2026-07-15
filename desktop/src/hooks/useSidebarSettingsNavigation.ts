import { useCallback, useState } from "react";
import type { MouseEvent } from "react";
import type { AppPage } from "../App";
import type { SettingsSection } from "../components/SettingsPage";

interface UseSidebarSettingsNavigationParams {
  onNavigate: (page: AppPage) => void;
  onOpenSettingsSection?: (section: SettingsSection) => void;
}

export function useSidebarSettingsNavigation({
  onNavigate,
  onOpenSettingsSection,
}: UseSidebarSettingsNavigationParams) {
  const [settingsMenuOpen, setSettingsMenuOpen] = useState(false);

  const closeSettingsMenu = useCallback(() => {
    setSettingsMenuOpen(false);
  }, []);

  const openSettingsSection = useCallback(
    (section: SettingsSection) => {
      setSettingsMenuOpen(false);
      if (onOpenSettingsSection) {
        onOpenSettingsSection(section);
      } else {
        onNavigate("settings");
      }
    },
    [onNavigate, onOpenSettingsSection],
  );

  const openGeneralSettings = useCallback(() => {
    openSettingsSection("general");
  }, [openSettingsSection]);

  const toggleSettingsMenu = useCallback((event: MouseEvent) => {
    event.stopPropagation();
    setSettingsMenuOpen((open) => !open);
  }, []);

  return {
    settingsMenuOpen,
    closeSettingsMenu,
    openSettingsSection,
    openGeneralSettings,
    toggleSettingsMenu,
  };
}
