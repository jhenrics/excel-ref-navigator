# Reference Navigator — Excel Add-in

An Excel add-in that shows all **precedent** (and dependent) cells for the currently selected cell, grouped by sheet, with one-click navigation — even across worksheets.

## Features

- **Scan precedents**: See every cell referenced by the active cell's formula
- **Scan dependents**: See every cell whose formula references the active cell
- **Direct vs indirect**: Badges show whether a reference is direct (first level) or indirect (chained)
- **Cross-sheet navigation**: Click any reference to jump to it, even on another sheet
- **Auto-refresh**: Optionally re-scans whenever you change your selection
- **Cell values shown**: Each reference displays its current value for quick context

## Project Structure

```
excel-ref-navigator/
├── manifest.xml          # Office Add-in manifest (registers the taskpane)
├── package.json          # Dependencies & scripts
├── webpack.config.js     # Build config (dev server on https://localhost:3000)
├── src/
│   ├── taskpane.html     # Taskpane UI (HTML + embedded CSS)
│   └── taskpane.js       # Core logic (Office JS API calls)
└── assets/               # Icons (you'll add placeholder PNGs here)
```

## Setup Instructions

### Prerequisites
- **Node.js** ≥ 18
- **Excel** (desktop on Windows/Mac, or Excel on the web)

### 1. Install dependencies
```bash
cd excel-ref-navigator
npm install
```

### 2. Create placeholder icons
The manifest references icon images. Create simple placeholder PNGs:
```bash
# Or use any 16x16, 32x32, 80x80 PNG images
# You can generate them with ImageMagick if installed:
convert -size 16x16 xc:#217346 assets/icon-16.png
convert -size 32x32 xc:#217346 assets/icon-32.png
convert -size 80x80 xc:#217346 assets/icon-80.png
```
Or simply place any PNG icons in the `assets/` folder with those names.

### 3. Start the dev server
```bash
npm run sideload
```
This installs HTTPS dev certs and starts webpack-dev-server on `https://localhost:3000`.

### 4. Sideload the add-in in Excel

#### Option A: Excel Desktop (Windows)
1. Open Excel
2. Go to **Insert → My Add-ins → Upload My Add-in**
3. Browse to `manifest.xml` and click OK
4. The "Ref Navigator" button appears on the Home tab

#### Option B: Excel Desktop (Mac)
1. Copy `manifest.xml` to `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/`
2. Restart Excel
3. The add-in appears on the Home tab

#### Option C: Excel on the Web
1. Go to excel.office.com and open a workbook
2. Go to **Insert → Office Add-ins → Upload My Add-in**
3. Browse to `manifest.xml`

### 5. Use the add-in
1. Click the **Ref Navigator** button on the Home ribbon
2. Select any cell with a formula
3. Click **Scan References** (or leave auto-refresh on)
4. Click any item in the list to jump to that cell

## Using Claude Code to Extend This

Here are some things you can ask Claude Code to help with:

- **"Add highlighting"** — Temporarily color precedent cells in the sheet
- **"Add a tree view"** — Show the full dependency chain as a collapsible tree
- **"Add formula parsing"** — Parse the formula string to show which part references which cell
- **"Support named ranges"** — Resolve named ranges to their cell addresses
- **"Add a search/filter box"** — Filter the reference list by sheet name or address
- **"Package for production"** — Build for AppSource submission

## Key Office JS APIs Used

| API | Purpose |
|-----|---------|
| `Range.getDirectPrecedents()` | Get cells directly referenced by the formula |
| `Range.getPrecedents()` | Get all precedents (full chain) |
| `Range.getDirectDependents()` | Get cells directly depending on this cell |
| `Range.getDependents()` | Get all dependents (full chain) |
| `Worksheet.activate()` | Switch to another sheet |
| `Range.select()` | Navigate to and select a cell |
| `Worksheet.onSelectionChanged` | Auto-refresh when selection changes |

These APIs require **ExcelApi 1.12+** (available in Excel 2021, Microsoft 365, and Excel on the web).

## Troubleshooting

- **"ItemNotFound" error**: The cell has no precedents/dependents — this is handled gracefully in the UI
- **Dev server cert errors**: Run `npx office-addin-dev-certs install` to trust the local HTTPS certificate
- **Add-in not loading**: Make sure the dev server is running on `https://localhost:3000` and the manifest URLs match
