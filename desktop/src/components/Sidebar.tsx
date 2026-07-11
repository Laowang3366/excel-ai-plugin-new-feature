import React, { useEffect, useCallback, useMemo, useRef } from "react";
import { useChatStore } from "../store/chatStore";
import { useSettingsStore } from "../store/settingsStore";
import type { AppPage } from "../App";
import type { SettingsSection } from "./SettingsPage";
import { getAppText } from "../i18n";
import { useExcelConnection } from "../hooks/useExcelConnection";
import { useOfficeConnection } from "../hooks/useOfficeConnection";
import { useDocumentDismiss } from "../hooks/useDocumentDismiss";
import { useSidebarFolderFiles } from "../hooks/useSidebarFolderFiles";
import { useSidebarResize } from "../hooks/useSidebarResize";
import { useSidebarSectionToggles } from "../hooks/useSidebarSectionToggles";
import { useSidebarSortMenu } from "../hooks/useSidebarSortMenu";
import { useSidebarSettingsNavigation } from "../hooks/useSidebarSettingsNavigation";
import { useSidebarThreadContextMenu } from "../hooks/useSidebarThreadContextMenu";
import { useSidebarThreadCreation } from "../hooks/useSidebarThreadCreation";
import { useSidebarViewedThreads } from "../hooks/useSidebarViewedThreads";
import { HostSelectionDialog } from "./excel/HostSelectionDialog";
import { SidebarSearchPalette } from "./sidebar/SidebarSearchPalette";
import { SidebarCollapsed } from "./sidebar/SidebarCollapsed";
import { buildSidebarDerivedLists } from "../utils/sidebarHelpers";
import type {
  FolderSectionActions,
  FolderSectionFileMenuApi,
  FolderSectionThreadActions,
} from "./sidebar/FolderSection";
import { SidebarExpanded } from "./sidebar/SidebarExpanded";

export type { IntentKind } from "../utils/sidebarHelpers";
interface SidebarProps {
  collapsed: boolean;
  currentPage: AppPage;
  onNavigate: (page: AppPage) => void;
  onOpenSettingsSection?: (section: SettingsSection) => void;
}

export const Sidebar: React.FC<SidebarProps> = ({
  collapsed,
  currentPage,
  onNavigate,
  onOpenSettingsSection,
}) => {
  const {
    threads,
    activeThreadId,
    runningThreadIds,
    turnStatus,
    loadThreads,
    switchThread,
    createNewThread,
    deleteThread,
    moveThreadToFolder,
    addFilesToComposer,
  } = useChatStore();
  const { language, pinnedFolders, addPinnedFolder, removePinnedFolder, updatePinnedFolder } =
    useSettingsStore();
  const text = getAppText(language);

  const sidebarRef = useRef<HTMLElement>(null);
  const {
    searchOpen,
    projectsExpanded,
    conversationsExpanded,
    toggleSearch,
    closeSearch,
    toggleProjectsExpanded,
    toggleConversationsExpanded,
  } = useSidebarSectionToggles();
  const { sidebarWidth, isResizing, handleResizeStart } = useSidebarResize(288);
  const { viewedThreadStatusAt, markThreadViewed } = useSidebarViewedThreads(
    threads,
    activeThreadId,
  );
  const {
    creatingFolderThread,
    creatingNewThread,
    handleCreateFolderThread,
    handleCreateNewThread,
  } = useSidebarThreadCreation(createNewThread);
  const {
    settingsMenuOpen,
    closeSettingsMenu,
    openSettingsSection,
    openGeneralSettings,
    toggleSettingsMenu,
  } = useSidebarSettingsNavigation({ onNavigate, onOpenSettingsSection });
  const {
    contextMenu,
    openThreadContextMenu,
    handleConfirmDelete,
    handleMoveToFolder,
    closeContextMenu,
    setContextMoveMenu,
    setContextConfirming,
  } = useSidebarThreadContextMenu({ deleteThread, moveThreadToFolder });

  const {
    excelStatus,
    connecting,
    connectFailed,
    pulseDot,
    pendingHosts,
    handleConnect,
    handleSelectHost,
    setPendingHosts,
  } = useExcelConnection();
  const { wordStatus, presentationStatus } = useOfficeConnection();

  const {
    folderFiles,
    expandedFolders,
    fileContextMenu,
    handleAddFolder,
    handleToggleFolder,
    handleAddFile,
    handleFileContextMenu,
    closeFileContextMenu,
    handleTrashFile,
    handleOpenFile,
    handleCopyPath,
    handleRevealInExplorer,
    handlePinFile,
  } = useSidebarFolderFiles({
    pinnedFolders,
    searchOpen,
    addPinnedFolder,
    updatePinnedFolder,
    addFilesToComposer,
    onOpenFileMenu: closeContextMenu,
  });
  const {
    sortMenu,
    projectSortMode,
    conversationSortMode,
    handleOpenSortMenu,
    handleSelectSortMode,
    closeSortMenu,
  } = useSidebarSortMenu(() => {
    closeContextMenu();
    closeFileContextMenu();
  });

  useEffect(() => {
    loadThreads();
  }, [loadThreads]);

  const handleSwitchThread = useCallback(
    (threadId: string) => {
      markThreadViewed(threadId);
      switchThread(threadId);
    },
    [markThreadViewed, switchThread],
  );

  const handleThreadContextMenu = useCallback(
    (e: React.MouseEvent, threadId: string, inFolder?: boolean) => {
      closeFileContextMenu();
      openThreadContextMenu(e, threadId, inFolder);
    },
    [closeFileContextMenu, openThreadContextMenu],
  );

  const folderActions = useMemo<FolderSectionActions>(
    () => ({
      toggle: handleToggleFolder,
      createThread: handleCreateFolderThread,
      remove: removePinnedFolder,
    }),
    [handleCreateFolderThread, handleToggleFolder, removePinnedFolder],
  );

  const folderThreadActions = useMemo<FolderSectionThreadActions>(
    () => ({
      switchThread: handleSwitchThread,
      openContextMenu: handleThreadContextMenu,
    }),
    [handleSwitchThread, handleThreadContextMenu],
  );

  const folderFileMenuApi = useMemo<FolderSectionFileMenuApi>(
    () => ({
      state: fileContextMenu,
      addFile: handleAddFile,
      openContextMenu: handleFileContextMenu,
      close: closeFileContextMenu,
      trashFile: handleTrashFile,
      openFile: handleOpenFile,
      copyPath: handleCopyPath,
      revealInExplorer: handleRevealInExplorer,
      pinFile: handlePinFile,
    }),
    [
      closeFileContextMenu,
      fileContextMenu,
      handleAddFile,
      handleCopyPath,
      handleFileContextMenu,
      handleOpenFile,
      handlePinFile,
      handleRevealInExplorer,
      handleTrashFile,
    ],
  );

  useDocumentDismiss({ active: contextMenu !== null, onDismiss: closeContextMenu });
  useDocumentDismiss({ active: fileContextMenu !== null, onDismiss: closeFileContextMenu });
  useDocumentDismiss({ active: settingsMenuOpen, onDismiss: closeSettingsMenu });
  useDocumentDismiss({ active: sortMenu !== null, onDismiss: closeSortMenu });

  const dismissPendingHosts = useCallback(() => {
    setPendingHosts(null);
  }, [setPendingHosts]);

  const hasSearchQuery = false;
  const {
    ungroupedThreads,
    groupedByFolder,
    hasProjectItems,
    hasConversationItems,
    showNoSearchResults,
  } = useMemo(
    () =>
      buildSidebarDerivedLists({
        threads,
        pinnedFolders,
        folderFiles,
        projectSortMode,
        conversationSortMode,
        language,
        hasSearchQuery,
      }),
    [
      conversationSortMode,
      folderFiles,
      hasSearchQuery,
      language,
      pinnedFolders,
      projectSortMode,
      threads,
    ],
  );

  const hostSelectionDialog =
    pendingHosts && pendingHosts.length > 1 ? (
      <HostSelectionDialog
        availableHosts={pendingHosts}
        onSelect={handleSelectHost}
        onDismiss={dismissPendingHosts}
      />
    ) : null;

  const sidebarContent = collapsed ? (
    <SidebarCollapsed
      currentPage={currentPage}
      text={text}
      creatingNewThread={creatingNewThread}
      excelStatus={excelStatus}
      wordStatus={wordStatus}
      presentationStatus={presentationStatus}
      onCreateNewThread={handleCreateNewThread}
      onToggleSearch={toggleSearch}
      onAddFolder={handleAddFolder}
      onOpenSettings={openGeneralSettings}
    />
  ) : (
    <SidebarExpanded
      currentPage={currentPage}
      text={text}
      language={language}
      sidebarRef={sidebarRef}
      isResizing={isResizing}
      sidebarWidth={sidebarWidth}
      projectsExpanded={projectsExpanded}
      conversationsExpanded={conversationsExpanded}
      groupedByFolder={groupedByFolder}
      ungroupedThreads={ungroupedThreads}
      hasSearchQuery={hasSearchQuery}
      hasProjectItems={hasProjectItems}
      hasConversationItems={hasConversationItems}
      showNoSearchResults={showNoSearchResults}
      activeThreadId={activeThreadId}
      runningThreadIds={runningThreadIds}
      turnStatus={turnStatus}
      creatingNewThread={creatingNewThread}
      creatingFolderThread={creatingFolderThread}
      viewedThreadStatusAt={viewedThreadStatusAt}
      expandedFolders={expandedFolders}
      contextMenu={contextMenu}
      pinnedFolders={pinnedFolders}
      sortMenu={sortMenu}
      projectSortMode={projectSortMode}
      conversationSortMode={conversationSortMode}
      settingsMenuOpen={settingsMenuOpen}
      searchOpen={searchOpen}
      excelStatus={excelStatus}
      wordStatus={wordStatus}
      presentationStatus={presentationStatus}
      connectFailed={connectFailed}
      connecting={connecting}
      pulseDot={pulseDot}
      onResizeStart={handleResizeStart}
      onCreateNewThread={handleCreateNewThread}
      onToggleSearch={toggleSearch}
      onToggleProjectsExpanded={toggleProjectsExpanded}
      onToggleConversationsExpanded={toggleConversationsExpanded}
      onOpenSortMenu={handleOpenSortMenu}
      onAddFolder={handleAddFolder}
      folderActions={folderActions}
      folderFileMenuApi={folderFileMenuApi}
      folderThreadActions={folderThreadActions}
      onSwitchThread={handleSwitchThread}
      onThreadContextMenu={handleThreadContextMenu}
      onSelectSortMode={handleSelectSortMode}
      onConfirmDelete={handleConfirmDelete}
      onMoveToFolder={handleMoveToFolder}
      onPinThread={closeContextMenu}
      onRenameThread={closeContextMenu}
      onCloseContextMenu={closeContextMenu}
      onSetContextMoveMenu={setContextMoveMenu}
      onSetContextConfirming={setContextConfirming}
      onConnect={handleConnect}
      onToggleSettingsMenu={toggleSettingsMenu}
      onOpenSettingsSection={openSettingsSection}
      onCloseSettingsMenu={closeSettingsMenu}
    />
  );

  return (
    <>
      {sidebarContent}
      <SidebarSearchPalette
        open={searchOpen}
        threads={threads}
        folders={pinnedFolders}
        folderFiles={folderFiles}
        language={language}
        activeThreadId={activeThreadId}
        onClose={closeSearch}
        onSwitchThread={handleSwitchThread}
        onAddFile={handleAddFile}
        onCreateNewThread={handleCreateNewThread}
        onAddFolder={handleAddFolder}
        onOpenSettings={openGeneralSettings}
      />
      {hostSelectionDialog}
    </>
  );
};
