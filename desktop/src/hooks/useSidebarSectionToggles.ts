import { useCallback, useState } from "react";

export function useSidebarSectionToggles() {
  const [searchOpen, setSearchOpen] = useState(false);
  const [projectsExpanded, setProjectsExpanded] = useState(true);
  const [conversationsExpanded, setConversationsExpanded] = useState(true);

  const toggleSearch = useCallback(() => {
    setSearchOpen((open) => !open);
  }, []);

  const closeSearch = useCallback(() => {
    setSearchOpen(false);
  }, []);

  const toggleProjectsExpanded = useCallback(() => {
    setProjectsExpanded((expanded) => !expanded);
  }, []);

  const toggleConversationsExpanded = useCallback(() => {
    setConversationsExpanded((expanded) => !expanded);
  }, []);

  return {
    searchOpen,
    projectsExpanded,
    conversationsExpanded,
    toggleSearch,
    closeSearch,
    toggleProjectsExpanded,
    toggleConversationsExpanded,
  };
}
