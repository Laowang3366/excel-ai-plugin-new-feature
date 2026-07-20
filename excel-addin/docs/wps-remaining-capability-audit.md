# WPS remaining capability audit (codex/excel-addin)

**Baseline:** `a70ad2e` era review.
**Scope:** remaining `wps-jsa` **unsupported** rows after CF/DV, structure, format, formula governance.
**Evidence order:** `desktop/public/wps-jsa-bridge` + in-repo WPS member-probe modules > desktop COM Worker (not a JSA contract) > guesswork (forbidden).

## In-repo WPS JSA surface (verified)

`desktop/public/wps-jsa-bridge/main.js` only exercises:

- `window.Application` / `ActiveWorkbook` / `Name`
- `JSIDE.SelectedJSComponent.CodeModule` (desktop localhost bridge for macro write)

It does **not** document ListObjects, ChartObjects, Shapes, FreezePanes, PageSetup, DisplayGridlines, Range.Locked, spill APIs, or `getImage`.

Load-item runtime continues to use **member probes** on assumed ET COM-shaped objects (`Range`, `Worksheets`, `Names`, `FormatConditions`, `Validation`, …) without claiming official JSA.

## Remaining unsupported (do not enlarge without new evidence)

| Capability | Why still unsupported | Desktop-only evidence (not sufficient alone) |
|---|---|---|
| `range.read` expand spill / currentArray | No WPS spill member contract in bridge or add-in probes | Office.js only |
| `formula.protection.*` | No in-repo `Range.Locked` / formula-cell lock probe for WPS | Office.js `format.protection.locked` |
| table / unlist / filter / sort | No ListObjects contract | Desktop table actions |
| chart 全系 / chart.image | No ChartObjects/setData contract | Office.js charts |
| shape / range.image | No Shapes / Range.getImage | Office.js |
| `sheet.freeze.*` | Needs `ActiveWindow.FreezePanes` or sheet freeze panes; **not** in JSA bridge; only desktop `ExcelTemplatePrintActionService` COM | COM Worker freeze |
| `sheet.display.*` | `ActiveWindow.DisplayGridlines` etc. desktop COM only | COM Worker display |
| `sheet.pageLayout.*` | `PageSetup` desktop COM print path only | COM Worker PageSetup |
| macros / PQ / Pivot / OpenXML / workbook open-create-save-switch | Product boundary | desktop |

**Decision this round:** **no new WPS feature implementation**. Forging freeze/display/pageLayout/table/chart would require guessing host members beyond bridge + existing probes.

## Package delivery boundary (text vs runtime)

| Kind | Example | Gate |
|---|---|---|
| **Documentation / prompt text** | `sourcePath: "desktop/electron/..."`, Chinese boundary “禁止 Electron/COM/child_process” | **Allowed** in source and shipped bundles |
| **Build-time Node CLI** | `package-prod.mjs` / `package-wps-jsa.mjs` `import { spawnSync } from "node:child_process"` | **Allowed** in scripts only; excluded from source runtime scan; **must not** appear inside task-pane `dist` JS |
| **Runtime import/require** | `require("electron")`, `from "child_process"`, `Wengge.OfficeWorker`, Office Interop | **Forbidden** in add-in source (except excluded package CLIs) and **forbidden** in packaged `dist` text artifacts |

Implementation:

- Shared detector: `excel-addin/scripts/runtimeDesktopDeps.mjs`
- Source tests: `excel-addin/tests/noDesktopDeps.test.ts` (text vs runtime fixtures)
- Package-time assert on WPS + Office production packages after layout is written

## Next candidates (only if evidence appears)

1. Documented WPS JSA sample or real sideload log showing `ActiveWindow.FreezePanes` / `Range.Locked` / `ListObjects` from **in-process task pane** (not desktop Worker).
2. Official WPS JSAPI docs for the same members used by the add-in host path.

Until then keep matrix **implemented\*** only for probed paths and **unsupported** for the table above.
