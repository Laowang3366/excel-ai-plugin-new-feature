import React from "react";
import type { AppPage } from "../../App";
import type { ThreadMetadata, FolderFileInfo } from "../../electronApi";
import type { AppLanguage, PinnedFolder } from "../../store/settingsStore";
import type { ExcelStatus } from "../../utils/sidebarHelpers";
import type { OfficeAppStatus } from "../../hooks/useOfficeConnection";
import type { getAppText } from "../../i18n";
import type { SettingsSection } from "../SettingsPage";
import { FolderSection, UngroupedThreadList } from "./FolderSection";
import { ThreadContextMenu, type ContextMenuState } from "./ThreadContextMenu";
import type { FileContextMenuState } from "./FileContextMenu";
import { SidebarFooter } from "./SidebarFooter";
import {
  Check,
  ChevronDown,
  ChevronRight,
  ClipboardList,
  Clock,
  PenLine,
  Plus,
  RefreshCw,
  Search,
} from "../common/IconMap";

export type SidebarSortSection = "projects" | "conversations";
export type SidebarSortMode = "recentDesc" | "recentAsc" | "nameAsc" | "nameDesc";

export interface SidebarGroupedFolder {
  folder: PinnedFolder;
  threads: ThreadMetadata[];
  files: FolderFileInfo[];
}

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
  fileContextMenu: FileContextMenuState | null;
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
  onToggleFolder: (folderPath: string) => void;
  onCreateFolderThread: (folderPath: string) => void;
  onRemoveFolder: (folderPath: string) => void;
  onAddFile: (file: FolderFileInfo) => void;
  onSwitchThread: (threadId: string) => void;
  onThreadContextMenu: (event: React.MouseEvent, threadId: string, inFolder?: boolean) => void;
  onFileContextMenu: (event: React.MouseEvent, file: FolderFileInfo, isPinned: boolean) => void;
  onFileContextMenuClose: () => void;
  onTrashFile: (filePath: string) => void;
  onOpenFile: (filePath: string) => void;
  onCopyPath: (filePath: string) => void;
  onRevealInExplorer: (filePath: string) => void;
  onPinFile: (filePath: string) => void;
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
  fileContextMenu,
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
  onToggleFolder,
  onCreateFolderThread,
  onRemoveFolder,
  onAddFile,
  onSwitchThread,
  onThreadContextMenu,
  onFileContextMenu,
  onFileContextMenuClose,
  onTrashFile,
  onOpenFile,
  onCopyPath,
  onRevealInExplorer,
  onPinFile,
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
    <aside className={`sidebar${isResizing ? " no-transition" : ""}`} ref={sidebarRef} style={{ width: isResizing || sidebarWidth !== 260 ? sidebarWidth : undefined }}>
      <div className={`sidebar-resize-handle${isResizing ? " resizing" : ""}`} onMouseDown={onResizeStart} />
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

      <div className="sidebar-content">
        <div className="sidebar-section-group">
          <div className="sidebar-section-header">
            <button className="sidebar-section-toggle" onClick={onToggleProjectsExpanded}>
              <span>{text.sidebar.projects}</span>
              {projectsExpanded ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
            </button>
            <button className="sidebar-section-add" onClick={(e) => onOpenSortMenu(e, "projects")} title={text.sidebar.sort}>
              <ClipboardList size={14} />
            </button>
            <button className="sidebar-section-add" onClick={onAddFolder} title={text.sidebar.addFolder}>
              <Plus size={14} />
            </button>
          </div>
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
                  onToggleFolder={onToggleFolder}
                  onCreateFolderThread={onCreateFolderThread}
                  onRemoveFolder={onRemoveFolder}
                  onAddFile={onAddFile}
                  onSwitchThread={onSwitchThread}
                  onThreadContextMenu={onThreadContextMenu}
                  fileContextMenu={fileContextMenu}
                  onFileContextMenu={onFileContextMenu}
                  onFileContextMenuClose={onFileContextMenuClose}
                  onTrashFile={onTrashFile}
                  onOpenFile={onOpenFile}
                  onCopyPath={onCopyPath}
                  onRevealInExplorer={onRevealInExplorer}
                  onPinFile={onPinFile}
                />
              ))}
              {groupedByFolder.length === 0 && !hasSearchQuery && (
                <div className="sidebar-section-empty">{text.sidebar.noProjects}</div>
              )}
            </div>
          )}
        </div>

        <div className="sidebar-section-group">
          <div className="sidebar-section-header">
            <button className="sidebar-section-toggle" onClick={onToggleConversationsExpanded}>
              <span>{text.sidebar.conversations}</span>
              {conversationsExpanded ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
            </button>
            <button className="sidebar-section-add" onClick={(e) => onOpenSortMenu(e, "conversations")} title={text.sidebar.sort}>
              <ClipboardList size={14} />
            </button>
            <button
              className={`sidebar-section-add${creatingNewThread ? " creating" : ""}`}
              onClick={onCreateNewThread}
              disabled={creatingNewThread}
              title={text.sidebar.newThread}
            >
              {creatingNewThread ? <RefreshCw size={14} className="spin" /> : <Plus size={14} />}
            </button>
          </div>
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
        ) : (!hasSearchQuery && !hasProjectItems && !hasConversationItems && (
          <div className="sidebar-empty">{text.sidebar.noThreads}</div>
        ))}
      </div>

      {sortMenu && (
        <div
          className="sidebar-sort-menu"
          style={{ left: sortMenu.x, top: sortMenu.y }}
          data-section={sortMenu.section}
          onClick={(e) => e.stopPropagation()}
        >
          {([
            ["recentDesc", text.sidebar.sortRecentDesc],
            ["recentAsc", text.sidebar.sortRecentAsc],
            ["nameAsc", text.sidebar.sortNameAsc],
            ["nameDesc", text.sidebar.sortNameDesc],
          ] as const).map(([mode, label]) => {
            const activeMode = sortMenu.section === "projects" ? projectSortMode : conversationSortMode;
            return (
              <button
                key={mode}
                className={`sidebar-sort-menu-item${activeMode === mode ? " active" : ""}`}
                onClick={() => onSelectSortMode(sortMenu.section, mode)}
              >
                <Clock size={14} />
                <span>{label}</span>
                {activeMode === mode && <Check size={14} />}
              </button>
            );
          })}
        </div>
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
