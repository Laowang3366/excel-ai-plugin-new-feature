# Open Access And P0 Fixes Design

## Goal

Make the desktop application available without activation or license checks, remove the activation server and its data, and fix the two remaining P0 Office issues: approval bypass through safe tools and destructive Excel style replacement.

## Scope

### Remove activation and licensing

- Delete the standalone `admin-server` directory, including its source, tests, UI, activation database, WAL/SHM files, and logs.
- Remove desktop activation state management, heartbeat handling, offline tolerance, activation IPC handlers, preload APIs, renderer stores, activation dialogs, administration views, device management views, and navigation entries.
- Remove startup and feature gates that depend on activation status. The desktop application must open directly into its normal working UI.
- Remove activation-only types, tests, imports, documentation, and build references.
- Do not retain a feature flag or compatibility mode. Existing activation files on previously installed user machines are ignored because no runtime code will read them.

### Prevent Office approval bypass

- Treat `office.action.inspect` and `office.action.validate` as strictly read-only tools.
- Define explicit operation compatibility rules instead of trusting the caller-provided `action`.
- Reject mutation operations at the executor boundary before the Office bridge is called.
- Repeat the same compatibility validation inside the Office action adapter so direct internal calls cannot bypass the executor check.
- Keep `office.action.apply` as the approved entry point for mutation operations.

The regression contract is that operations such as `writeRange`, `setDataValidation`, `setHeaderFooter`, `createPresentation`, and `addSlides` cannot execute through either safe tool.

### Preserve existing Excel styles

- When `xl/styles.xml` already exists, preserve all existing style nodes and their ordering.
- Append the new table-header font, fill, and cell format to the existing `fonts`, `fills`, and `cellXfs` collections.
- Update each collection's `count` attribute and use the appended `cellXfs` index for styled header cells.
- Preserve number formats, borders, named styles, differential formats, and all existing style indices.
- Create the current minimal style sheet only when the workbook has no `xl/styles.xml`.
- Keep workbook content-type and relationship handling for newly created style parts.

Repeated styling may append another style definition, but it must never invalidate or renumber existing styles. Deduplication is outside this P0 repair.

## Architecture

The activation removal is a deletion-oriented change: callers and UI gates are removed first, followed by the now-unreachable activation modules and server. No replacement authorization abstraction is introduced.

Office operation authorization is enforced at two existing boundaries. A shared pure helper classifies whether an action kind may execute an operation. The executor uses it for user-facing rejection, while the adapter uses it as a defensive invariant.

Excel style preservation remains inside the existing OpenXML table styling module. The implementation calculates appended indices from the current style collections and returns the new cell style index to the worksheet styling functions.

## Error Handling

- Safe Office tools return a failed tool result with a clear message when given a mutation operation, and must not invoke any file bridge.
- If an existing Excel style sheet lacks a required collection or has an invalid count, the implementation derives the count from child elements and inserts the missing collection without replacing unrelated XML.
- If the style sheet cannot be modified safely, the operation fails without writing the output file.

## Testing

### Activation removal

- Type checking and production build prove that activation imports, IPC contracts, and UI references have been removed.
- A repository search must find no runtime references to activation APIs, activation stores, license keys, heartbeat handling, or `admin-server`.
- Existing non-activation desktop workflows remain covered by the full test suite.

### Office authorization

- Executor tests call both safe tools with representative Excel, Word, and presentation mutation operations and assert rejection.
- Adapter tests call mutation operations with `action: "inspect"` and `action: "validate"` and assert a failed result without invoking the implementation bridge.
- Existing legitimate inspect and validate operations continue to pass.

### Excel styles

- A workbook fixture contains existing date and currency number formats, custom fonts and fills, borders, named styles, and cells using multiple style indices.
- After table styling, the fixture's original XML fragments and indices remain present.
- The new header cells use the newly appended style index rather than `s="1"`.
- The existing no-style workbook test continues to verify creation of a valid minimal style sheet.

## Verification

Run:

1. Targeted red-green tests for Office authorization.
2. Targeted red-green tests for Excel style preservation.
3. Desktop lint and TypeScript checks.
4. The full desktop Vitest suite.
5. The desktop production build.
6. A final repository search for activation and `admin-server` references.

The work is complete when the desktop starts without activation code, the server and its data are absent, safe tools cannot route mutations, existing Excel styles survive table styling, and all verification commands pass apart from any explicitly documented pre-existing warnings.
