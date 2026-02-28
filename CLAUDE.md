# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is
An Excel Office Add-in (taskpane) called **Reference Navigator** that lets users see all cells referenced by (or referencing) the active cell's formula, and click to navigate to them — even across worksheets.

## Tech Stack
- **Office JS API** (ExcelApi 1.11+, 1.12 for indirect) — `getDirectPrecedents()`, `getPrecedents()`, `getDependents()`, `getDirectDependents()`
- **Vanilla JS** — no framework, single `taskpane.js` entry point
- **Webpack 5** — bundles and serves via `webpack-dev-server` on `https://localhost:3000`
- **manifest.xml** — standard Office Add-in XML manifest

## Commands
- `npm run dev` — Start dev server with hot reload (or double-click `start.bat`)
- `npm run build` — Production build to `dist/`
- `npm run sideload` — Install HTTPS dev certs + start dev server
- `npm run lint` — Run ESLint on `src/`
- `npm run validate` — Validate manifest.xml
- `start-hidden.vbs` — Runs dev server silently (for Windows startup)

No test framework is configured.

## Key Files
- `src/taskpane.js` — All logic: scanning, rendering, navigation, formula parsing (~580 lines)
- `src/taskpane.html` — UI + embedded CSS (single file, no separate stylesheet)
- `manifest.xml` — Add-in registration, ribbon button, permissions
- `webpack.config.js` — Build config (uses Office dev certs for HTTPS)

## Architecture Decisions
- CSS is embedded in the HTML for simplicity (single taskpane file)
- Direct precedents from `getDirectPrecedents()` shown as main list, sorted by formula order (parsed from formula string via regex)
- Indirect precedents from `getPrecedents()` shown separately under "Indirect" header (these APIs return merged ranges, so direct/indirect are queried separately to avoid mismatches)
- **Merged range expansion**: `getDirectPrecedents()` merges adjacent individually-referenced cells into ranges (e.g. `$B$65` and `$B$66` become `B65:B66`). The code cross-references with the parsed formula to split them back into separate entries when they aren't referenced as an explicit range. See `expandMergedRanges()`.
- **Resolved formula bar**: after scanning precedents, a bar shows the formula with cell references replaced by their display values (e.g. `=$B$66*$B$65` → `=18.0*50.0%`). Range references are kept as-is. See `buildResolvedFormula()`.
- **Navigation mode**: clicking a result navigates to it but suppresses auto-refresh so the list stays stable; a "Back to source" button returns to the original cell
- Auto-refresh uses workbook-level `onSelectionChanged` (works across sheets)
- Navigation uses `Worksheet.activate()` + `Range.select()` for cross-sheet jumps
- Add-in is sideloaded via Windows registry (`HKCU\Software\Microsoft\Office\16.0\WEF\Developer`)
- Sheet name shown inline on refs only when they're on a different sheet

## Constraints
- ExcelApi 1.11+ required for direct precedents/dependents
- ExcelApi 1.12+ required for indirect (all) precedents/dependents
- These APIs don't work across workbook boundaries
- `getPrecedents()` merges adjacent cells into ranges — don't use for direct/indirect classification
- Large dependency chains (1000+ cells) may be slow
