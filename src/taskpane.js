/* ================================================================
   Reference Navigator — Excel Add-in
   Shows precedent/dependent cells for the active cell and lets
   the user click to navigate to any of them (even across sheets).
   ================================================================ */

Office.onReady(({ host }) => {
  if (host === Office.HostType.Excel) {
    init();
  }
});

// ── State ──────────────────────────────────────────────────────
let mode = "precedents"; // "precedents" | "dependents"
let selectionHandler = null;
let isNavigating = false; // true while stepping through results
let sourceCell = null; // { address, sheetName } of the scanned cell

// ── DOM refs ───────────────────────────────────────────────────
function $(id) {
  return document.getElementById(id);
}

// ── Initialise ─────────────────────────────────────────────────
async function init() {
  $("btnScan").addEventListener("click", () => {
    exitNavMode();
    scan();
  });
  $("btnClear").addEventListener("click", () => {
    exitNavMode();
    clearResults();
  });

  $("togglePrecedents").addEventListener("click", () => setMode("precedents"));
  $("toggleDependents").addEventListener("click", () => setMode("dependents"));

  $("autoRefresh").addEventListener("change", toggleAutoRefresh);

  // Check API support
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.11")) {
    $("results").innerHTML = `
      <div class="error-msg">
        <strong>Error:</strong> This add-in requires Excel API 1.11+.
        Please update Excel to use Reference Navigator.
      </div>`;
    return;
  }

  // Initial read of active cell (don't scan yet — wait for user action or selection change)
  await updateActiveCell();

  // Register selection-changed listener (auto-refresh on by default)
  registerSelectionHandler();
}

// ── Mode toggle ────────────────────────────────────────────────
function setMode(newMode) {
  mode = newMode;
  $("togglePrecedents").classList.toggle("active", mode === "precedents");
  $("toggleDependents").classList.toggle("active", mode === "dependents");
  exitNavMode();
  scan();
}

// ── Navigation mode ────────────────────────────────────────────
function exitNavMode() {
  isNavigating = false;
  sourceCell = null;
  const btn = $("btnReturn");
  if (btn) btn.remove();
}

async function returnToSource() {
  if (!sourceCell) return;
  const src = sourceCell;
  exitNavMode();
  try {
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(src.sheetName);
      ws.activate();
      const range = ws.getRange(src.address);
      range.select();
      await ctx.sync();
    });
  } catch (err) {
    console.error("Return failed:", err);
  }
}

// ── Auto-refresh ───────────────────────────────────────────────
function toggleAutoRefresh() {
  if ($("autoRefresh").checked) {
    registerSelectionHandler();
  } else {
    unregisterSelectionHandler();
  }
}

function registerSelectionHandler() {
  if (selectionHandler) return;
  Excel.run(async (ctx) => {
    selectionHandler = ctx.workbook.onSelectionChanged.add(onSelectionChanged);
    await ctx.sync();
  }).catch(() => {});
}

function unregisterSelectionHandler() {
  // We can't easily remove cross-sheet, so we just flag it off
  selectionHandler = null;
}

async function onSelectionChanged() {
  // If we're stepping through results, don't re-scan
  if (isNavigating) return;
  await updateActiveCell();
  scan();
}

// ── Update the active-cell bar ─────────────────────────────────
async function updateActiveCell() {
  try {
    await Excel.run(async (ctx) => {
      const cell = ctx.workbook.getActiveCell();
      cell.load(["address", "formulas"]);
      await ctx.sync();

      const displayAddr = cell.address.includes("!")
        ? cell.address.split("!")[1]
        : cell.address;
      $("cellAddress").textContent = displayAddr;
      const formula = cell.formulas[0][0];
      if (typeof formula === "string" && formula.startsWith("=")) {
        $("cellFormula").textContent = formula;
        $("cellFormula").title = formula;
      } else {
        $("cellFormula").textContent = "No formula";
        $("cellFormula").title = "";
      }
    });
  } catch {
    $("cellAddress").textContent = "—";
    $("cellFormula").textContent = "";
  }
}

// ── Main scan ──────────────────────────────────────────────────
async function scan() {
  const results = $("results");
  $("resolvedFormula").style.display = "none";
  results.innerHTML = `
    <div class="loading">
      <div class="spinner"></div>
      <p>Scanning ${mode}…</p>
    </div>`;

  await updateActiveCell();

  try {
    await Excel.run(async (ctx) => {
      const cell = ctx.workbook.getActiveCell();
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      cell.load(["address", "formulas"]);
      sheet.load("name");
      await ctx.sync();

      // Remember which cell we scanned so we can return to it
      sourceCell = {
        address: cell.address.includes("!")
          ? cell.address.split("!")[1]
          : cell.address,
        sheetName: sheet.name,
      };

      const formula = cell.formulas[0][0];
      const hasFormula =
        typeof formula === "string" && formula.startsWith("=");

      // For precedents the cell must have a formula;
      // for dependents any cell can have dependents.
      if (mode === "precedents" && !hasFormula) {
        results.innerHTML = "";
        return;
      }

      // Get DIRECT + ALL (to label them)
      let directAreas, allAreas;
      const hasAllApi = Office.context.requirements.isSetSupported("ExcelApi", "1.12");
      try {
        if (mode === "precedents") {
          directAreas = cell.getDirectPrecedents();
          allAreas = hasAllApi ? cell.getPrecedents() : null;
        } else {
          directAreas = cell.getDirectDependents();
          allAreas = hasAllApi ? cell.getDependents() : null;
        }
        directAreas.areas.load("address");
        if (allAreas) allAreas.areas.load("address");
        await ctx.sync();
      } catch (e) {
        if (
          e instanceof OfficeExtension.Error &&
          (e.code === "ItemNotFound" || e.code === "GeneralException")
        ) {
          results.innerHTML = `
            <div class="state-msg">
              <svg width="36" height="36" viewBox="0 0 24 24" fill="none"
                stroke="#bbb" stroke-width="1.5" stroke-linecap="round">
                <circle cx="11" cy="11" r="8"/>
                <line x1="21" y1="21" x2="16.65" y2="16.65"/>
              </svg>
              <p>No ${mode} found for <strong>${cell.address}</strong></p>
              <p class="hint">This cell doesn't reference${mode === "precedents" ? "" : " or isn't referenced by"} any other cells.</p>
            </div>`;
          return;
        }
        throw e;
      }
      // If all-API not available, use direct as the full set
      if (!allAreas) allAreas = directAreas;

      // Each RangeAreas.address may contain comma-separated ranges
      function parseAddresses(areasCollection) {
        const addrs = [];
        for (const rangeAreas of areasCollection.items) {
          for (const part of rangeAreas.address.split(",")) {
            addrs.push(part.trim());
          }
        }
        return addrs;
      }

      function parseAddr(addr) {
        const bangIdx = addr.lastIndexOf("!");
        let sheetName, cellRef;
        if (bangIdx > 0) {
          sheetName = addr.substring(0, bangIdx).replace(/^'|'$/g, "");
          cellRef = addr.substring(bangIdx + 1);
        } else {
          sheetName = sourceCell.sheetName;
          cellRef = addr;
        }
        const isOtherSheet = sheetName !== sourceCell.sheetName;
        return {
          displayAddress: isOtherSheet ? sheetName + "!" + cellRef : cellRef,
          cellRef,
          fullAddress: addr,
          sheetName,
        };
      }

      // Build direct and indirect lists separately
      const directAddrs = parseAddresses(directAreas.areas);
      const allAddrs = allAreas ? parseAddresses(allAreas.areas) : [];

      // Indirect items — from getPrecedents, excluding exact matches with direct
      const directSet = new Set(directAddrs);
      const indirectItems = allAddrs
        .filter((a) => !directSet.has(a))
        .map((a) => ({ ...parseAddr(a), isDirect: false }));

      // Expand merged API ranges back into individual cells when the formula
      // references them individually (API merges adjacent single-cell refs)
      const formulaOrder = extractFormulaRefs(formula, sourceCell.sheetName);
      const expandedDirectAddrs = expandMergedRanges(
        directAddrs, formulaOrder, sourceCell.sheetName
      );

      // Direct items — from getDirectPrecedents (with merged ranges expanded)
      const directItems = expandedDirectAddrs.map((a) => ({ ...parseAddr(a), isDirect: true }));

      // Sort direct refs by order of appearance in the formula
      directItems.sort((a, b) => {
        const aN = a.sheetName + "!" + a.cellRef.replace(/\$/g, "");
        const bN = b.sheetName + "!" + b.cellRef.replace(/\$/g, "");
        const aIdx = formulaOrder.indexOf(aN);
        const bIdx = formulaOrder.indexOf(bN);
        return (aIdx < 0 ? 99999 : aIdx) - (bIdx < 0 ? 99999 : bIdx);
      });

      const allRefs = [...directItems, ...indirectItems];

      // Load values for each referenced range
      const refRanges = [];
      for (const item of allRefs) {
        const range = ctx.workbook.worksheets
          .getItem(item.sheetName)
          .getRange(item.cellRef);
        range.load(["values", "text", "address"]);
        refRanges.push({ range, item });
      }
      await ctx.sync();

      // Attach display values
      for (const { range, item } of refRanges) {
        const row = range.text && range.text[0];
        const val = row ? row[0] : null;
        item.displayValue =
          val !== undefined && val !== null && val !== "" ? String(val) : "(empty)";
      }

      // Build resolved formula (refs replaced with values)
      if (mode === "precedents" && hasFormula) {
        const resolved = buildResolvedFormula(formula, directItems, sourceCell.sheetName);
        const el = $("resolvedFormula");
        if (resolved && resolved !== formula) {
          el.textContent = resolved;
          el.title = resolved;
          el.style.display = "block";
        } else {
          el.style.display = "none";
        }
      } else {
        $("resolvedFormula").style.display = "none";
      }

      // ── Render ───────────────────────────────
      let html = `
        <div class="summary-bar">
          <span>${allRefs.length} ${mode} found for ${escHtml(sourceCell.sheetName)}!${escHtml(sourceCell.address)}</span>
        </div>`;

      function renderRefList(items) {
        let out = `<ul class="ref-list">`;
        for (const item of items) {
          out += `
            <li class="ref-item" data-address="${escAttr(item.fullAddress)}" data-sheet="${escAttr(item.sheetName)}">
              <div class="ref-info">
                <div class="ref-address">${escHtml(item.displayAddress)}</div>
                <div class="ref-value">${escHtml(item.displayValue)}</div>
              </div>
              <span class="ref-arrow">→</span>
            </li>`;
        }
        out += `</ul>`;
        return out;
      }

      html += renderRefList(directItems);
      if (indirectItems.length > 0) {
        html += `<div class="sheet-group-header">Indirect</div>`;
        html += renderRefList(indirectItems);
      }

      results.innerHTML = html;

      // Attach click-to-navigate handlers
      document.querySelectorAll(".ref-item").forEach((el) => {
        el.addEventListener("click", () => {
          // Enter navigation mode — suppress auto-refresh
          isNavigating = true;

          // Highlight active item
          document.querySelectorAll(".ref-item.active").forEach((e) =>
            e.classList.remove("active")
          );
          el.classList.add("active");

          // Show return button if not already shown
          showReturnButton();

          const addr = el.dataset.address;
          const sh = el.dataset.sheet;
          navigateTo(sh, addr);
        });
      });
    });
  } catch (err) {
    results.innerHTML = `
      <div class="error-msg">
        <strong>Error:</strong> ${escHtml(err.message || String(err))}
      </div>`;
  }
}

// ── Show/hide the "Return to source" button ─────────────────────
function showReturnButton() {
  if ($("btnReturn")) return;
  const btn = document.createElement("button");
  btn.id = "btnReturn";
  btn.className = "btn btn-return";
  btn.textContent = `← Back to ${sourceCell.sheetName}!${sourceCell.address}`;
  btn.addEventListener("click", returnToSource);
  // Insert before results
  const results = $("results");
  results.parentNode.insertBefore(btn, results);
}

// ── Navigate to a cell (even on another sheet) ─────────────────
async function navigateTo(sheetName, fullAddress) {
  try {
    await Excel.run(async (ctx) => {
      let range;
      if (sheetName && sheetName !== "(current sheet)") {
        const ws = ctx.workbook.worksheets.getItem(sheetName);
        ws.activate();
        const cellRef = fullAddress.includes("!")
          ? fullAddress.split("!")[1]
          : fullAddress;
        range = ws.getRange(cellRef);
      } else {
        range = ctx.workbook.worksheets
          .getActiveWorksheet()
          .getRange(fullAddress);
      }
      range.select();
      await ctx.sync();
    });
  } catch (err) {
    console.error("Navigation failed:", err);
  }
}

// ── Clear ──────────────────────────────────────────────────────
function clearResults() {
  $("resolvedFormula").style.display = "none";
  $("results").innerHTML = `
    <div class="state-msg">
      <svg width="36" height="36" viewBox="0 0 24 24" fill="none"
        stroke="#bbb" stroke-width="1.5" stroke-linecap="round">
        <path d="M9 18l6-6-6-6"/>
      </svg>
      <p>Select a cell and click <strong>Scan References</strong></p>
      <p class="hint">or enable auto-refresh above.</p>
    </div>`;
}

// ── Extract cell references from a formula in order ────────────
function extractFormulaRefs(formula, currentSheet) {
  if (!formula || typeof formula !== "string") return [];
  // Match references like: 'Sheet Name'!$A$1:$B$2, Sheet1!A1, A1:B2, $C$3
  const refPattern = /(?:'[^']+'|[A-Za-z0-9_]+)!\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?|\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?/g;
  const refs = [];
  let match;
  while ((match = refPattern.exec(formula)) !== null) {
    const raw = match[0];
    const bangIdx = raw.lastIndexOf("!");
    let sheet, cellRef;
    if (bangIdx > 0) {
      sheet = raw.substring(0, bangIdx).replace(/^'|'$/g, "");
      cellRef = raw.substring(bangIdx + 1);
    } else {
      sheet = currentSheet;
      cellRef = raw;
    }
    // Normalize: strip $ signs
    refs.push(sheet + "!" + cellRef.replace(/\$/g, ""));
  }
  return refs;
}

// ── Build resolved formula (refs → values) ─────────────────────
// Replaces single-cell references in the formula with their display
// values.  Range references (e.g. A1:A10) are kept as-is.
function buildResolvedFormula(formula, refItems, currentSheet) {
  if (!formula || !formula.startsWith("=")) return null;

  // Lookup: "Sheet!Cell" (normalized, no $) → displayValue
  const valueMap = new Map();
  for (const item of refItems) {
    const key = item.sheetName + "!" + item.cellRef.replace(/\$/g, "");
    valueMap.set(key, item.displayValue);
  }

  const refPattern =
    /(?:'[^']+'|[A-Za-z0-9_]+)!\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?|\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?/g;

  return formula.replace(refPattern, (match) => {
    const bangIdx = match.lastIndexOf("!");
    const cellPart = bangIdx > 0 ? match.substring(bangIdx + 1) : match;

    // Skip range references
    if (cellPart.includes(":")) return match;

    let sheet, cellRef;
    if (bangIdx > 0) {
      sheet = match.substring(0, bangIdx).replace(/^'|'$/g, "");
      cellRef = match.substring(bangIdx + 1);
    } else {
      sheet = currentSheet;
      cellRef = match;
    }
    const key = sheet + "!" + cellRef.replace(/\$/g, "");
    return valueMap.has(key) ? valueMap.get(key) : match;
  });
}

// ── Expand merged API ranges ────────────────────────────────────
// getDirectPrecedents() merges adjacent individually-referenced cells
// into ranges (e.g. B65,B66 → B65:B66). This splits them back when
// the formula references them as separate cells, not as a range.
function expandMergedRanges(addresses, formulaRefs, currentSheet) {
  const expanded = [];
  for (const addr of addresses) {
    const bangIdx = addr.lastIndexOf("!");
    let sheetPrefix = "", sheetName, cellRef;
    if (bangIdx > 0) {
      sheetPrefix = addr.substring(0, bangIdx + 1);
      sheetName = addr.substring(0, bangIdx).replace(/^'|'$/g, "");
      cellRef = addr.substring(bangIdx + 1);
    } else {
      sheetName = currentSheet;
      cellRef = addr;
    }

    // Single cell — keep as-is
    if (!cellRef.includes(":")) {
      expanded.push(addr);
      continue;
    }

    // Check if this exact range appears in the formula (e.g. SUM(B65:B66))
    const normalizedRange = sheetName + "!" + cellRef.replace(/\$/g, "");
    if (formulaRefs.includes(normalizedRange)) {
      expanded.push(addr);
      continue;
    }

    // Expand range into individual cells, keep those found in formula
    const cells = expandRangeToCells(cellRef);
    const matched = [];
    for (const cell of cells) {
      if (formulaRefs.includes(sheetName + "!" + cell)) {
        matched.push(sheetPrefix + cell);
      }
    }

    if (matched.length > 0) {
      expanded.push(...matched);
    } else {
      // Formula parsing missed these refs — keep original range
      expanded.push(addr);
    }
  }
  return expanded;
}

function expandRangeToCells(cellRef) {
  const [startRef, endRef] = cellRef.split(":");
  const startCol = startRef.replace(/[\$\d]/g, "");
  const startRow = parseInt(startRef.replace(/[\$A-Z]/gi, ""), 10);
  const endCol = endRef.replace(/[\$\d]/g, "");
  const endRow = parseInt(endRef.replace(/[\$A-Z]/gi, ""), 10);

  const sc = colLetterToNum(startCol), ec = colLetterToNum(endCol);
  const cells = [];
  for (let r = startRow; r <= endRow; r++) {
    for (let c = sc; c <= ec; c++) {
      cells.push(colNumToLetter(c) + r);
    }
  }
  return cells;
}

function colLetterToNum(col) {
  let n = 0;
  for (let i = 0; i < col.length; i++) {
    n = n * 26 + (col.toUpperCase().charCodeAt(i) - 64);
  }
  return n;
}

function colNumToLetter(num) {
  let col = "";
  while (num > 0) {
    const mod = (num - 1) % 26;
    col = String.fromCharCode(65 + mod) + col;
    num = Math.floor((num - 1) / 26);
  }
  return col;
}

// ── Helpers ────────────────────────────────────────────────────
function escHtml(s) {
  const d = document.createElement("div");
  d.textContent = s;
  return d.innerHTML;
}
function escAttr(s) {
  return s.replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}
