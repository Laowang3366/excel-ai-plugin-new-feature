# OpenXML Office Engine Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a COM-free Open XML editing path for Office files and show structured edit details in a right-side preview panel.

**Architecture:** Keep COM bridges as live-application adapters and add an independent file-editing engine under `agent/tools/implementations/officeOpenXml`. File tools return structured edit summaries, and the frontend projects those summaries into a right-side monitor panel from completed tool results.

**Tech Stack:** TypeScript, Vitest, Electron IPC tool events, ZIP/XML processing via direct project dependencies.

---

### Task 1: Open XML File Engine Slice

**Files:**
- Create: `desktop/electron/agent/tools/implementations/officeOpenXml/types.ts`
- Create: `desktop/electron/agent/tools/implementations/officeOpenXml/officeOpenXmlEngine.ts`
- Create: `desktop/electron/agent/tools/implementations/officeOpenXml/officeOpenXmlEngine.test.ts`
- Modify: `desktop/package.json`

- [x] Write failing tests for inspecting and replacing text in minimal `.docx`, `.pptx`, and `.xlsx` packages.
- [x] Run `npm test -- electron/agent/tools/implementations/officeOpenXml/officeOpenXmlEngine.test.ts` and confirm missing engine failure.
- [x] Implement the smallest ZIP/XML engine that reads document parts, returns summaries, replaces text, and writes a target file.
- [x] Re-run the targeted test and confirm it passes.

### Task 2: Agent Tool Wiring

**Files:**
- Modify: `desktop/electron/agent/tools/contracts/office.ts`
- Modify: `desktop/electron/agent/tools/registry/office.ts`
- Modify: `desktop/electron/agent/tools/executors/officeExecutors.ts`
- Modify: `desktop/electron/agent/tools/executors/createToolExecutors.ts`
- Modify: `desktop/electron/agent/runtime/agentRuntime.ts`
- Modify: `desktop/electron/agent/runtime/bridgeRegistry.ts`
- Modify: `desktop/electron/agent/tools/registry/officeTools.test.ts`
- Modify: `desktop/electron/agent/prompts/sections/officeToolsPrompt.ts`

- [x] Write failing tests that require `office.file.inspect` and `office.file.replaceText`.
- [x] Run the Office tool test and confirm the new tools are missing.
- [x] Add the file bridge contract, registry definitions, executor wiring, runtime injection, and prompt guidance.
- [x] Re-run the Office tool test and confirm it passes.

### Task 3: Right-Side Edit Monitor

**Files:**
- Create: `desktop/src/utils/officeEditEvents.ts`
- Create: `desktop/src/components/office/OfficePreviewPanel.tsx`
- Modify: `desktop/src/components/ChatPage.tsx`
- Modify: `desktop/src/styles/app-layout.css`
- Modify: `desktop/src/styles/chat.css`
- Test: `desktop/src/utils/officeEditEvents.test.ts`

- [x] Write failing tests that extract file edit summaries from completed tool results.
- [x] Run the frontend utility test and confirm the extractor is missing.
- [x] Implement the extractor and a right-side panel that shows current file, operation timeline, replacement counts, and output path.
- [x] Re-run the utility test and confirm it passes.

### Task 4: Review, Verify, Commit

**Files:**
- All files above.

- [x] Run `npm test`.
- [x] Run `npm run typecheck`.
- [x] Run `npm run build`.
- [x] Run a smoke test using generated temporary `.docx` and `.pptx` files, then delete temporary files.
- [x] Review the diff for layer boundaries and unrelated files.
- [x] Commit only the reviewed Open XML and preview-panel files.
