# Excel parity audit (add-in vs desktop)

Phase60 evidence close-out + parity classification.

**Phase:** Phase60
**Baseline HEAD:** `c46362f81c2a7c334ab0ded60e854de287fedd12` (`c46362f8`)
**Scope:** documentation / evidence only — no runtime, desktop, or installer code changes.
**Evidence order:** running source under `excel-addin/shared/**` and `desktop/electron/agent/tools/**` > phase history text.

## 1. Real-device WPS evidence (narrow, closed)

| Fact | Value |
|---|---|
| Host | WPS 12.1.0.26885 |
| Package gitSha | `c46362f8` |
| Install | `wps:status` `current=true`, `drift=[]`, addon dir `WenggeExcelAiAddin_` |
| Ribbon | Same install state; **cold start** restored 「文格 AI」 tab; click 「打开助手」 → task pane full render; **no** package/code change → prior absence = load/cache **transient**, not code regression |
| `selection.get` | Blank workbook, Sheet1, selection G17 → `ok:true`, `tool:"selection.get"`, `sheetName:"Sheet1"`, `address:"G17"`, `values:[[null]]` |

**Do not expand:** this is **not** a pass for other WPS tools, Office.js Excel sideload, or full 98-tool device matrix.
`implemented*` elsewhere remains **member-probe + mock/unit only**.

Sources for the closed Address path:

- `excel-addin/shared/host/wpsJsaSelection.ts`
- `excel-addin/shared/host/wpsJsaAddress.ts`
- `excel-addin/tests/wpsJsaAddress.test.ts`

### 1.1 Task-pane layout viewport mismatch (device-measured, not a tool pass)

| Fact | Value |
|---|---|
| Win32 child | WPS task pane CEF: `CefBrowserWindow` / `Chrome_RenderWidgetHostHWND` **Left=1402, Width=646, Height=906** |
| Packaged HTML | Install `index.html` retains viewport meta (unchanged packaging contract) |
| Chromium control | Ordinary Chromium **300×1000**: UI full single-column, **no clip** |
| WPS full-screen shot | Content begins ~**+370px** from task-pane left edge; **right side clipped** |
| Playwright repro | Viewport **1428px** width: `.app { max-width:720px; margin:0 auto }` → left margin `(1428-720)/2 = 354px` — matches WPS visual offset |

**Root cause (evidence-backed):** WPS CEF reports a **~1428px layout viewport** while the **visible child is only ~646px** wide. Centered `.app` (`excel-addin/src/styles.css`) shifts content into the clipped region.

**Not claimed:** full WPS tool pass; Excel sideload; that Office.js hosts share this CEF quirk.

**Host kind signal already in code (no UA):** `detectHostKind()` → `"wps-jsa"` (`excel-addin/shared/host/detectHost.ts`); `App.tsx` already holds `hostKind` state / badge.

## 2. Inventory sources (counts are not equivalence)

| Surface | Count / shape | Source files |
|---|---|---|
| Add-in `TOOL_DEFINITIONS` | **98** tools | `excel-addin/shared/tools/definitions.ts` + `*Definitions.ts` |
| Desktop **window** Excel tools | workbook / range / formula / sheet / macro / ui registries | `desktop/electron/agent/tools/registry/{workbook,range,formula,sheet,macro,ui}.ts` |
| Desktop Excel **operations** | 35 `EXCEL_CAPABILITIES` ops (via `office.action.*`) | `desktop/electron/agent/tools/officeCore/excelCapabilities.ts` |
| Desktop Office orchestration | workflow / transaction / documents | `desktop/electron/agent/tools/registry/office.ts`, `officeReliability.ts` |

Tool-name count ≠ feature parity. Desktop often maps many ops into fewer public tools (`office.action.apply`) plus COM/OpenXML engines the add-in deliberately lacks.

## 3. Desktop window tools → add-in classification

### 3.1 Current-active-workbook — add-in implemented (Office.js + WPS path)

| Desktop tool | Add-in tool(s) | Office.js | WPS | Evidence |
|---|---|---|---|---|
| `selection.get` | `selection.get` | implemented | **device-verified** @ G17 (`implemented*` elsewhere still member-probe) | `wpsJsaSelection.ts`, `officeJsAdapter` selection |
| `range.read` / `write` / `clear` | same | implemented | implemented* | `wpsJsaRangeRead.ts`, `wpsJsaAdapter.ts` |
| `formula.context` | `formula.context` (+ `formula.read`/`write`) | implemented | implemented* | `wpsJsaAdapter.getFormulaContext` |
| `sheet.operation` | `sheet.operation` + `sheet.list/add/rename/delete` | implemented | implemented* (copy/move in `wpsJsaSheetOps.ts`) | `sheetOperation.ts`, `wpsJsaSheetOps.ts` |
| `workbook.inspect` | `workbook.inspect` | implemented | implemented* | `wpsJsaInspect.ts` |
| `workbook.save` | `workbook.save` | implemented | implemented* (`ActiveWorkbook.Save` probe) | `wpsJsaWorkbookSave.ts` (bridge evidence in `desktop/public/wps-jsa-bridge/main.js` text only; **no** device pass claimed) |

### 3.2 Product boundary / host API unsupported in add-in (do not implement here)

| Desktop tool | Why out of add-in scope | Desktop source |
|---|---|---|
| `workbook.open` / `workbook.create` / `workbook.switch` | Disk path + process document lifecycle; needs COM/OpenXML/Worker | `registry/workbook.ts` |
| `macro.detect` / `macro.run` / `macro.write` | VBA / WPS JSA IDE; desktop localhost bridge `JSIDE.CodeModule` | `registry/macro.ts`, `desktop/public/wps-jsa-bridge/main.js` |
| `ui.addControl` / `removeControl` / `listControls` / `createForm` / `addMenu` | Excel forms / ActiveX-style UI via COM | `registry/ui.ts` |

### 3.3 Office.js-implemented in add-in but **no WPS contract**

These exist as add-in tools for Excel Online/Desktop Office.js; WPS returns typed `unsupported` (see `wpsJsaUnsupported.ts`, `wpsJsa*Unsupported.ts`):

- Tables: `table.*`, filter/sort/unlist (`wpsJsaUnsupported.ts`)
- Charts + series + axes + dataLabels + image (`wpsJsaUnsupported.ts`, `wpsJsaChartSeriesUnsupported.ts`, Office.js under `officeJsChart*`)
- Shapes: `shape.*` (`officeJsShapes.ts` / WPS unsupported)
- Freeze / display / pageLayout (`officeJsFreeze.ts`, `officeJsSheetDisplay.ts`, `officeJsPageLayout.ts` — WPS no bridge contract)
- `range.image.get` / `chart.image.get` (`officeJsRangeImage.ts`, `officeJsChartImage.ts`)
- `formula.protection.*` (`officeJsFormulaProtection*.ts` — no WPS `Range.Locked` contract)
- Pivot / slicer (`officeJsPivot*.ts`, `officeJsSlicer*.ts` / `wpsJsaPivotUnsupported.ts`, `wpsJsaSlicerUnsupported.ts`)
- Workbook templates (`officeJsTemplate*.ts` / `wpsJsaTemplateUnsupported.ts`)

## 4. Desktop `EXCEL_CAPABILITIES` (35 ops) → add-in

Source: `desktop/electron/agent/tools/officeCore/excelCapabilities.ts`.

| Desktop operation | preferredEngine | Add-in mapping | Class |
|---|---|---|---|
| `createWorkbook` | openxml | none (`workbook.create` desktop-only) | **product boundary** (disk/OpenXML) |
| `insertChart` | openxml→com | `chart.create` (+ series tools) | Office.js implemented; WPS unsupported |
| `applyConditionalFormatting` | openxml→com | `conditionalFormat.*` | active workbook Office.js + WPS* |
| `setDataValidation` | openxml→com | `dataValidation.*` | active workbook Office.js + WPS* |
| `styleTable` | openxml | `table.create`/`table.update` | Office.js; WPS unsupported |
| `snapshot` | openxml→com | `range.image.get` / `chart.image.get` (memory Base64 only) | Office.js; WPS unsupported; **no** path write |
| `createPivotTable` / `refreshPivotTables` | com | `pivot.create` / `pivot.refresh` | Office.js; WPS unsupported |
| `addSlicer` | com | `slicer.*` | Office.js; WPS unsupported |
| `createPowerQuery` / `inspectPowerQueries` / `managePowerQuery` | com | **none** | **product boundary** (COM PQ) |
| `inspectCharts` / `formatChart` | com | `chart.*` / series / axes / markers / trendlines | Office.js subset; WPS unsupported |
| `inspectWorkbookObjects` / `manageWorkbookObject` / `manageWorksheetObjects` | com | `workbook.objects.inspect` + object update tools | Office.js partial; WPS sheets/names* only (`wpsJsaWorkbookObjects.ts`) |
| `captureWorkbookTemplate` / `applyWorkbookTemplate` / `inspectWorkbookFormatting` | com | `workbook.template.capture` / `apply` | Office.js; WPS unsupported |
| `inspectPrintSettings` / `configurePrint` | com | `sheet.pageLayout.get/set` | Office.js; WPS unsupported |
| `exportSheetsToPdf` / `exportPdf` | com | **none** | **product boundary** (path/PDF COM) |
| `traceFormulaDependencies` / `inspectFormulaDependencies` / `repairFormulaReferences` / `convertFormulasToValues` / `inspectFormulaBackups` / `restoreFormulas` | com | `formula.dependencies.inspect` / `references.repair` / `convertToValues` / `backups.*` | active workbook pure-core + host read/write; WPS* |
| `inspectFormulaProtection` / `manageFormulaProtection` | com | `formula.protection.*` | Office.js; WPS unsupported |
| `exportRangeToWord` / `exportRangeToPresentation` / `buildReportPackage` | com | **none** | **product boundary** (cross Word/PPT/PDF) |

## 5. Add-in 98 `TOOL_DEFINITIONS` — buckets

Full names are enumerated by `name:` fields under `excel-addin/shared/tools/**` (count gate in tests: 98).

### A. Current-active-workbook — Office.js implemented; WPS implemented* (member-probe) except noted

`host.status`, `selection.get` (**device-verified G17**), `range.read|write|clear`, `range.format.read|write`, `range.insert|delete|autofit`, `formula.read|write|context`, formula governance (`dependencies.inspect`, `references.repair`, `convertToValues`, `backups.inspect|restore`), `sheet.list|add|rename|delete|operation`, `sheet.visibility.*`, `sheet.protection.*`, `namedRange.*`, `conditionalFormat.*`, `dataValidation.*`, `workbook.inspect`, `workbook.objects.inspect` (sheets/names partial on WPS), `workbook.save`.

Primary sources: `wpsJsaAdapter.ts`, `wpsJsaRangeRead.ts`, `wpsJsaFormat.ts`, `wpsJsaSheetOps.ts`, `wpsJsaStructure.ts`, `wpsJsaConditionalFormat.ts`, `wpsJsaDataValidation.ts`, `wpsJsaFormulaGovernance*.ts`, `wpsJsaNamedRanges.ts`, `wpsJsaWorkbookSave.ts`.

### B. Office.js implemented; WPS typed unsupported (no in-repo JSA contract)

`table.*` (list/create/delete/update/unlist/filter/sort), full `chart.*` tree, `shape.*`, `sheet.freeze.*`, `sheet.display.*`, `sheet.pageLayout.*`, `range.image.get`, `chart.image.get`, `formula.protection.*`, `pivot.*`, `slicer.*`, `workbook.template.*`.

Primary sources: `wpsJsaUnsupported.ts`, `wpsJsa*Unsupported.ts`, Office.js `officeJs*.ts`.

### C. Product / host boundaries (not add-in goals)

Disk open/create/saveAs/switch; macros; Power Query; COM/.NET/Electron/child_process; OpenXML file package ops; PDF/path export; cross-Office Word/PPT; persistent `office.workflow.*` / `office.transaction.*` (desktop `officeCore/workflow*.ts`, `transaction*.ts`).

### D. Real gaps still inside pure Office.js / WPS JSA (evidence-gated)

There is **no additional WPS feature** with both (1) in-repo bridge/member evidence and (2) missing implementation, after CF/DV/structure/format/governance batches.
See `excel-addin/docs/wps-remaining-capability-audit.md`.

Candidates that would require **new** host evidence before coding:

1. In-process task-pane log or official JSA sample for `ListObjects` / `ChartObjects` / `ActiveWindow.FreezePanes` / `PageSetup` / `Range.Locked`.
2. Official WPS JSAPI docs matching those members.

Until then: do **not** invent WPS members.

## 6. Next batch recommendation (minimal, contract-safe)

### 6.1 Priority: WPS task-pane layout fix (has real-device + Playwright evidence)

Trusted **UI** gap (not a new JSA member, not Electron/COM/.NET/child_process, no UA sniffing, no new deps):

1. When `hostKind === "wps-jsa"` (`detectHost.ts` / `App.tsx`), set a **stable** root class or `data-host="wps-jsa"` (not User-Agent).
2. WPS-only CSS: left-align `.app` (override `margin: 0 auto`), constrain layout width to the **visible** task-pane range (~646px evidence; avoid relying on CEF’s inflated ~1428 CSS viewport alone).
3. Tabs: `flex-wrap` + child `min-width: 0` so narrow panes do not force horizontal overflow.
4. Gates for that follow-up PR (out of Phase60 doc-only scope):
   - unit/DOM tests asserting WPS host attribute/class and left-aligned rules;
   - Playwright: viewport **1428** with left **646** clip window — content must stay visible (no 354px empty lead-in);
   - `package:wps` → install → WPS real-device screenshot re-check.
5. Source of the bug today: `excel-addin/src/styles.css` (`.app { max-width: 720px; margin: 0 auto; }`).

### 6.2 After layout: device / chat e2e (still no invented JSA members)

**No trusted WPS host-API code gap** remains for new table/chart/freeze/pageLayout/PQ/macro members without bridge/official contracts (see §5.D / `wps-remaining-capability-audit.md`).

Then:

1. **Real-device acceptance matrix** for existing `implemented*` paths (range read/write, sheet ops, format, CF/DV, formula governance, workbook.save) — one tool at a time; host logs only.
2. **Provider-chat end-to-end** on WPS task pane (model → tool loop → `selection.get` / `range.write`) with current 98 tools.
3. Optional: Microsoft Excel HTTPS sideload acceptance (still **unverified**).

Do **not** open table/chart/freeze/pageLayout/PQ/macro workstreams for WPS without new contracts. Do **not** “fix” layout via UA strings or new runtime dependencies.

## 7. Honesty rules (locked by tests)

- May document: install current, Ribbon cold-start restore, `selection.get` G17 payload @ `c46362f8`, task-pane CEF viewport vs visible-width mismatch (1428 vs 646) and centered-`.app` clip.
- Must **not** claim: all WPS capabilities device-passed; Excel sideload passed; `implemented*` = real sideload; layout issue “fixed” until a follow-up PR + device shot.
- Must **not** reintroduce “Ribbon / selection.get 待复验” for the closed G17 evidence without new regression facts.
