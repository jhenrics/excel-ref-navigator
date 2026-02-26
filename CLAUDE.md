# CLAUDE.md — Project Context for Claude Code

## What This Is
An Excel Office Add-in (taskpane) called **Reference Navigator** that lets users see all cells referenced by (or referencing) the active cell's formula, and click to navigate to them — even across worksheets.

## Tech Stack
- **Office JS API** (ExcelApi 1.12+) — `getDirectPrecedents()`, `getPrecedents()`, `getDependents()`, `getDirectDependents()`
- **Vanilla JS** — no framework, single `taskpane.js` entry point
- **Webpack 5** — bundles and serves via `webpack-dev-server` on `https://localhost:3000`
- **manifest.xml** — standard Office Add-in XML manifest

## Key Files
- `src/taskpane.js` — All logic: scanning, rendering, navigation
- `src/taskpane.html` — UI + embedded CSS (single file, no separate stylesheet)
- `manifest.xml` — Add-in registration, ribbon button, permissions
- `webpack.config.js` — Build config

## Architecture Decisions
- CSS is embedded in the HTML for simplicity (single taskpane file)
- Uses `WorkbookRangeAreas` from the precedents/dependents API, grouped by worksheet
- Navigation uses `Worksheet.activate()` + `Range.select()` for cross-sheet jumps
- Auto-refresh listens to `Worksheet.onSelectionChanged`

## Common Tasks
- `npm run sideload` — Start dev server + install certs
- `npm run build` — Production build to `dist/`
- Sideload `manifest.xml` in Excel via Insert → My Add-ins → Upload

## Constraints
- ExcelApi 1.12+ required (precedent/dependent APIs)
- These APIs don't work across workbook boundaries
- Large dependency chains (1000+ cells) may be slow
