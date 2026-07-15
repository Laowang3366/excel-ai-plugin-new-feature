import React from "react";
import type { AppPage } from "../../App";
import type { ThreadMetadata } from "../../electronApi";
import type { AppLanguage, PinnedFolder } from "../../store/settingsStore";
import type {
  ExcelStatus,
  SidebarGroupedFolder,
  SidebarSortMode,
} from "../../utils/sidebarHelpers";
import type { OfficeAppStatus } from "../../hooks/useOfficeConnection";
import type { getAppText } from "../../i18n";
import type { SettingsSection } from "../SettingsPage";
import {
  FolderSection,
  UngroupedThreadList,
  type FolderSectionActions,
  type FolderSectionFileMenuApi,
  type FolderSectionThreadActions,
} from "./FolderSection";
import { ThreadContextMenu, type ContextMenuState } from "./ThreadContextMenu";
import { SidebarFooter } from "./SidebarFooter";
import { Plus, RefreshCw } from "../common/IconMap";
import { SidebarSectionHeader } from "./SidebarSectionHeader";
import { SidebarSortMenu, type SidebarSortSection } from "./SidebarSortMenu";
import { SidebarExpandedToolbar } from "./SidebarExpandedToolbar";

export type { SidebarSortSection } from "./SidebarSortMenu";
export type { SidebarGroupedFolder, SidebarSortMode } from "../../utils/sidebarHelpers";

interface SidebarExpandedProps {
  currentPage: AppPage;
  text: ReturnType<typeof getAppText>;
  language: AppLanguage;
  sidebarRef: React.RefObject<HTMLElement>;
  isResizing: boolean;
  sidebarWidth: number;
  projectsExpanded: boolean;
  conversationsExpanded: boolean;
  groupedByFolder: SidebarGroupedFolder[];
  ungroupedThreads: ThreadMetadata[];
  hasSearchQuery: boolean;
  hasProjectItems: boolean;
  hasConversationItems: boolean;
  showNoSearchResults: boolean;
  activeThreadId: string | null;
  runningThreadIds: Record<string, boolean>;
  turnStatus: string;
  creatingNewThread: boolean;
  creatingFolderThread: string | null;
  viewedThreadStatusAt: Record<string, number>;
  expandedFolders: Record<string, boolean>;
  contextMenu: ContextMenuState | null;
  pinnedFolders: PinnedFolder[];
  sortMenu: { section: SidebarSortSection; x: number; y: number } | null;
  projectSortMode: SidebarSortMode;
  conversationSortMode: SidebarSortMode;
  settingsMenuOpen: boolean;
  searchOpen: boolean;
  excelStatus: ExcelStatus;
  wordStatus: OfficeAppStatus;
  presentationStatus: OfficeAppStatus;
  connectFailed: boolean;
  connecting: boolean;
  pulseDot: boolean;
  onResizeStart: (event: React.MouseEvent) => void;
  onCreateNewThread: () => void;
  onToggleSearch: () => void;
  onToggleProjectsExpanded: () => void;
  onToggleConversationsExpanded: () => void;
  onOpenSortMenu: (event: React.MouseEvent, section: SidebarSortSection) => void;
  onAddFolder: () => void;
  folderActions: FolderSectionActions;
  folderFileMenuApi: FolderSectionFileMenuApi;
  folderThreadActions: FolderSectionThreadActions;
  onSwitchThread: (threadId: string) => void;
  onThreadContextMenu: (event: React.MouseEvent, threadId: string, inFolder?: boolean) => void;
  onSelectSortMode: (section: SidebarSortSection, mode: SidebarSortMode) => void;
  onConfirmDelete: (threadId: string) => void;
  onMoveToFolder: (threadId: string, folderId?: string) => void;
  onPinThread: (threadId: string) => void;
  onRenameThread: (threadId: string) => void;
  onCloseContextMenu: () => void;
  onSetContextMoveMenu: (show: boolean) => void;
  onSetContextConfirming: (confirming: boolean) => void;
  onConnect: () => void;
  onToggleSettingsMenu: (event: React.MouseEvent) => void;
  onOpenSettingsSection: (section: SettingsSection) => void;
  onCloseSettingsMenu: () => void;
}

export function SidebarExpanded({
  currentPage,
  text,
  language,
  sidebarRef,
  isResizing,
  sidebarWidth,
  projectsExpanded,
  conversationsExpanded,
  groupedByFolder,
  ungroupedThreads,
  hasSearchQuery,
  hasProjectItems,
  hasConversationItems,
  showNoSearchResults,
  activeThreadId,
  runningThreadIds,
  turnStatus,
  creatingNewThread,
  creatingFolderThread,
  viewedThreadStatusAt,
  expandedFolders,
  contextMenu,
  pinnedFolders,
  sortMenu,
  projectSortMode,
  conversationSortMode,
  settingsMenuOpen,
  searchOpen,
  excelStatus,
  wordStatus,
  presentationStatus,
  connectFailed,
  connecting,
  pulseDot,
  onResizeStart,
  onCreateNewThread,
  onToggleSearch,
  onToggleProjectsExpanded,
  onToggleConversationsExpanded,
  onOpenSortMenu,
  onAddFolder,
  folderActions,
  folderFileMenuApi,
  folderThreadActions,
  onSwitchThread,
  onThreadContextMenu,
  onSelectSortMode,
  onConfirmDelete,
  onMoveToFolder,
  onPinThread,
  onRenameThread,
  onCloseContextMenu,
  onSetContextMoveMenu,
  onSetContextConfirming,
  onConnect,
  onToggleSettingsMenu,
  onOpenSettingsSection,
  onCloseSettingsMenu,
}: SidebarExpandedProps) {
  return (
    <aside
      className={`sidebar${isResizing ? " no-transition" : ""}`}
      ref={sidebarRef}
      style={{ width: isResizing || sidebarWidth !== 260 ? sidebarWidth : undefined }}
    >
      <div
        className={`sidebar-resize-handle${isResizing ? " resizing" : ""}`}
        onMouseDown={onResizeStart}
      />
      <SidebarExpandedToolbar
        text={text}
        creatingNewThread={creatingNewThread}
        searchOpen={searchOpen}
        onCreateNewThread={onCreateNewThread}
        onToggleSearch={onToggleSearch}
      />

      <div className="sidebar-content">
        <div className="sidebar-section-group">
          <SidebarSectionHeader
            title={text.sidebar.projects}
            expanded={projectsExpanded}
            sortTitle={text.sidebar.sort}
            actionTitle={text.sidebar.addFolder}
            actionIcon={<Plus size={14} />}
            onToggle={onToggleProjectsExpanded}
            onOpenSort={(event) => onOpenSortMenu(event, "projects")}
            onAction={onAddFolder}
          />
          {projectsExpanded && (
            <div className="sidebar-section-content">
              {groupedByFolder.map(({ folder, threads: folderThreads, files }) => (
                <FolderSection
                  key={folder.path}
                  folder={folder}
                  folderThreads={folderThreads}
                  files={files}
                  isExpanded={!!expandedFolders[folder.path]}
                  activeThreadId={activeThreadId}
                  runningThreadIds={runningThreadIds}
                  turnStatus={turnStatus}
                  creatingFolderThread={creatingFolderThread}
                  viewedThreadStatusAt={viewedThreadStatusAt}
                  language={language}
                  folderActions={folderActions}
                  threadActions={folderThreadActions}
                  fileMenuApi={folderFileMenuApi}
                />
              ))}
              {groupedByFolder.length === 0 && !hasSearchQuery && (
                <div className="sidebar-section-empty">{text.sidebar.noProjects}</div>
              )}
            </div>
          )}
        </div>

        <div className="sidebar-section-group">
          <SidebarSectionHeader
            title={text.sidebar.conversations}
            expanded={conversationsExpanded}
            sortTitle={text.sidebar.sort}
            actionTitle={text.sidebar.newThread}
            actionIcon={
              creatingNewThread ? <RefreshCw size={14} className="spin" /> : <Plus size={14} />
            }
            actionClassName={creatingNewThread ? " creating" : ""}
            actionDisabled={creatingNewThread}
            onToggle={onToggleConversationsExpanded}
            onOpenSort={(event) => onOpenSortMenu(event, "conversations")}
            onAction={onCreateNewThread}
          />
          {conversationsExpanded && (
            <div className="sidebar-section-content">
              <UngroupedThreadList
                threads={ungroupedThreads}
                activeThreadId={activeThreadId}
                runningThreadIds={runningThreadIds}
                turnStatus={turnStatus}
                viewedThreadStatusAt={viewedThreadStatusAt}
                language={language}
                onSwitchThread={onSwitchThread}
                onThreadContextMenu={onThreadContextMenu}
              />
            </div>
          )}
        </div>

        {showNoSearchResults ? (
          <div className="sidebar-empty">{text.sidebar.noSearchResults}</div>
        ) : (
          !hasSearchQuery &&
          !hasProjectItems &&
          !hasConversationItems && <div className="sidebar-empty">{text.sidebar.noThreads}</div>
        )}
      </div>

      {sortMenu && (
        <SidebarSortMenu
          menu={sortMenu}
          labels={{
            recentDesc: text.sidebar.sortRecentDesc,
            recentAsc: text.sidebar.sortRecentAsc,
            nameAsc: text.sidebar.sortNameAsc,
            nameDesc: text.sidebar.sortNameDesc,
          }}
          projectSortMode={projectSortMode}
          conversationSortMode={conversationSortMode}
          onSelectSortMode={onSelectSortMode}
        />
      )}
      {contextMenu && (
        <ThreadContextMenu
          contextMenu={contextMenu}
          pinnedFolders={pinnedFolders}
          language={language}
          onConfirmDelete={onConfirmDelete}
          onMoveToFolder={onMoveToFolder}
          onPinThread={onPinThread}
          onRenameThread={onRenameThread}
          onClose={onCloseContextMenu}
          setShowMoveMenu={onSetContextMoveMenu}
          setConfirming={onSetContextConfirming}
        />
      )}

      <SidebarFooter
        currentPage={currentPage}
        text={text}
        excelStatus={excelStatus}
        wordStatus={wordStatus}
        presentationStatus={presentationStatus}
        connectFailed={connectFailed}
        connecting={connecting}
        pulseDot={pulseDot}
        settingsMenuOpen={settingsMenuOpen}
        onConnect={onConnect}
        onToggleSettingsMenu={onToggleSettingsMenu}
        onOpenSettingsSection={onOpenSettingsSection}
        onCloseSettingsMenu={onCloseSettingsMenu}
      />
    </aside>
  );
}
