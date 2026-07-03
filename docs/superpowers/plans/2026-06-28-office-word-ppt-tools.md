# Word PPT Office Tools Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add first-class Word and PowerPoint tools so the agent can open, create, inspect, edit, and save documents and presentations like it already does for Excel.

**Architecture:** Keep Excel behavior unchanged and add separate Office COM bridges for Word and PowerPoint. Register new tool definitions and executors in the existing `toolRegistry` path, then instantiate the bridges in Electron main process alongside the Excel bridges.

**Tech Stack:** Electron main process, TypeScript, Windows COM automation through PowerShell, Vitest.

---

### Task 1: Tool Registry Contract

**Files:**
- Modify: `desktop/electron/agent/toolRegistry/interfaces.ts`
- Modify: `desktop/electron/agent/toolRegistry/definitions.ts`
- Modify: `desktop/electron/agent/toolRegistry/executors.ts`
- Test: `desktop/electron/agent/toolRegistry/officeTools.test.ts`

- [ ] Write failing tests that assert Word/PPT tool definitions exist and `createToolExecutors` forwards calls to fake Word/PPT/Office script bridges.
- [ ] Run `npm test -- electron/agent/toolRegistry/officeTools.test.ts` and confirm tests fail because tools are missing.
- [ ] Add Word/PPT bridge interfaces, safe/moderate/dangerous tool definitions, and executor wiring.
- [ ] Rerun the focused tests and confirm they pass.

### Task 2: COM Bridge Implementation

**Files:**
- Create: `desktop/electron/agent/officeBridge/wordComBridge.ts`
- Create: `desktop/electron/agent/officeBridge/presentationComBridge.ts`
- Create: `desktop/electron/agent/officeBridge/officeScriptBridge.ts`
- Create: `desktop/electron/agent/officeBridge/index.ts`
- Modify: `desktop/electron/agent/excelBridge/excelBridgeHelpers.ts` if reusable helper exports are needed.

- [ ] Add bridge methods for open/create/inspect/read/edit/save using PowerShell COM.
- [ ] Use Microsoft Office ProgIDs first and WPS candidate ProgIDs as fallback.
- [ ] Return clear errors when neither Office nor WPS COM is available.

### Task 3: Agent Integration

**Files:**
- Modify: `desktop/electron/main.ts`
- Modify: `desktop/electron/agent/systemPrompt.ts`
- Modify: `desktop/electron/main-modules/ipcHandlers.ts`

- [ ] Instantiate Word, Presentation, and Office script bridges in main.
- [ ] Pass the bridges to `createToolExecutors`.
- [ ] Add concise tool-selection guidance for Word/PPT editing.
- [ ] Update file dialogs and attachment guidance for `.doc`, `.docx`, `.ppt`, `.pptx`.

### Task 4: Verification

**Files:**
- Modify only files required by failing verification.

- [ ] Run focused Vitest suite for the new tool registry tests.
- [ ] Run `npm run typecheck`.
- [ ] Run a production build if typecheck passes quickly enough.
