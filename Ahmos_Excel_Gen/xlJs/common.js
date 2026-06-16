/* ═══════════════════════════════════════════════════════════
   common.js — Shared utilities  |  Ahmos Excel Formula Gen
   ═══════════════════════════════════════════════════════════ */
'use strict';

// ══════════════════════════════════════════════════════════
//  FILE PATH
// ══════════════════════════════════════════════════════════

/**
 * Validate and format a file path for Excel external references.
 * Handles full paths, UNC paths, filename-only, open/closed file distinction.
 * Returns { valid, formatted, isEmpty, isFullPath, error? }
 */
function processFilePath(rawPath) {
  if (!rawPath || !rawPath.trim()) return { valid: true, formatted: '', isEmpty: true };

  // Strip surrounding quotes
  let p = rawPath.trim().replace(/^["']|["']$/g, '').trim();

  const validExts = ['.xlsx', '.xlsm', '.xlsb', '.xls'];
  if (!validExts.some(e => p.toLowerCase().endsWith(e))) {
    return { valid: false, error: 'Invalid extension. Supported: .xlsx  .xlsm  .xlsb  .xls' };
  }

  // Find last directory separator
  let lastSlash = -1;
  for (let i = p.length - 1; i >= 0; i--) {
    if (p[i] === '\\' || p[i] === '/') { lastSlash = i; break; }
  }

  let formatted, isFullPath = false;
  if (lastSlash !== -1) {
    isFullPath = true;
    const dir  = p.substring(0, lastSlash + 1);
    const file = p.substring(lastSlash + 1);
    formatted = dir + '[' + file + ']';
  } else {
    formatted = '[' + p + ']';
  }

  return { valid: true, formatted, isEmpty: false, isFullPath };
}

/**
 * Build the full sheet-path prefix used inside Excel formula strings.
 * Handles sheet names that contain spaces or special characters by
 * always wrapping the entire workbook+sheet portion in single quotes.
 *
 * Examples:
 *   buildSheetPath('C:\\data\\file.xlsx', 'My Sheet')  →  "'C:\data\[file.xlsx]My Sheet'!"
 *   buildSheetPath('',                   'Sheet1')     →  "'Sheet1'!"
 *   buildSheetPath('',                   '')           →  ""
 */
function buildSheetPath(filePath, sheetName) {
  const fp = processFilePath(filePath);
  if (!fp.valid) return null;   // caller must handle
  const wb = fp.formatted;
  const sh = (sheetName || '').trim();
  if (!wb && !sh) return '';
  // Always quote: safe for spaces, hyphens, numbers-starting names, etc.
  if (!wb && sh)  return `'${sh}'!`;
  if (wb  && !sh) return `'${wb}'!`;
  return `'${wb}${sh}'!`;
}

// ══════════════════════════════════════════════════════════
//  INPUT NORMALISATION
// ══════════════════════════════════════════════════════════

/**
 * Normalise a raw range string typed by the user:
 *  - Trim whitespace
 *  - Uppercase
 *  - Remove trailing colon  "A2:"  → "A2"
 *  - Fix reversed colon-less col  "A2" stays "A2"
 * Returns the cleaned string ready for parseRange().
 */
function normaliseRangeInput(raw) {
  if (!raw) return '';
  let s = raw.trim().toUpperCase().replace(/\s/g, '');
  // Remove trailing colon
  s = s.replace(/:+$/, '');
  return s;
}

/**
 * Normalise a cell reference typed by the user.
 * Accepts e.g. "a2", " B3 ", "B$3" → "B3" (strips $ so it can be re-applied consistently)
 */
function normaliseCellRef(raw) {
  if (!raw) return '';
  return raw.trim().toUpperCase().replace(/\$/g, '').replace(/\s/g, '');
}

// ══════════════════════════════════════════════════════════
//  RANGE PARSER
// ══════════════════════════════════════════════════════════

/**
 * Parse a cleaned range string into components.
 * Handles: "B", "B:F", "B2", "B2:F", "B2:F200", "B:B"
 * Returns { startCol, startRow, endCol, endRow }
 *   startRow / endRow are numeric strings or '' if absent.
 */
function parseRange(input) {
  const s = normaliseRangeInput(input);
  const parts = s.split(':');
  const left  = parts[0] || '';
  const right = parts[1] !== undefined ? parts[1] : '';

  const startCol = left.replace(/\d/g, '')  || 'A';
  const startRow = left.replace(/[A-Z]/g, '') || '';
  const endCol   = right ? (right.replace(/\d/g, '') || startCol) : startCol;
  const endRow   = right ? right.replace(/[A-Z]/g, '') : '';

  return { startCol, startRow, endCol, endRow };
}

/**
 * Validate that a range string is parseable and makes structural sense.
 * Returns { ok, error? }
 */
function validateRange(input, label) {
  const s = normaliseRangeInput(input);
  if (!s) return { ok: false, error: `${label}: field is required.` };

  // Must start with at least one letter
  if (!/^[A-Z]/.test(s)) return { ok: false, error: `${label}: must start with a column letter (e.g., A or B2).` };

  const r = parseRange(s);

  // Column sanity: A–XFD
  const colToNum = c => c.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  if (colToNum(r.startCol) > 16384) return { ok: false, error: `${label}: column "${r.startCol}" is out of Excel range.` };
  if (colToNum(r.endCol)   > 16384) return { ok: false, error: `${label}: column "${r.endCol}" is out of Excel range.` };

  // If both rows given, start ≤ end
  if (r.startRow && r.endRow && parseInt(r.startRow) > parseInt(r.endRow)) {
    return { ok: false, error: `${label}: start row (${r.startRow}) is greater than end row (${r.endRow}).` };
  }

  return { ok: true };
}

/**
 * Check whether two ranges have mismatched explicit start rows and
 * return a warning message if they differ (after normalisation to defStart).
 */
function warnIfStartRowMismatch(retComp, tgtComp) {
  const r = parseInt(retComp.startRow || '1');
  const t = parseInt(tgtComp.startRow || '1');
  if (r !== t) {
    const used = Math.max(r, t);
    return `⚠️ Range start rows differ (Return: row ${r}, Filter: row ${t}). Both have been aligned to row ${used}. Please verify this is correct.`;
  }
  return null;
}

// ══════════════════════════════════════════════════════════
//  MATCH VALUE / EQUAL CELL
// ══════════════════════════════════════════════════════════

/**
 * Determine whether a match-value input is a cell reference or a literal string.
 * Cell ref:   A2, $B$3, AA10   → returns it normalised
 * Literal:    Sales, "Sales"   → wraps in Excel double-quotes
 *
 * Returns { type: 'cell'|'literal', formula: string }
 */
function resolveMatchValue(raw) {
  const s = raw.trim();
  // Strip surrounding quotes if user typed them
  const stripped = s.replace(/^["']|["']$/g, '').trim();

  // Cell reference pattern: optional $, letters, optional $, digits
  const cellPattern = /^\$?[A-Za-z]{1,3}\$?\d{1,7}$/;
  if (cellPattern.test(stripped)) {
    return { type: 'cell', formula: stripped.toUpperCase().replace(/\$/g, '') };
  }

  // Treat as literal string — escape any double-quotes inside
  const escaped = stripped.replace(/"/g, '""');
  return { type: 'literal', formula: `"${escaped}"` };
}

// ══════════════════════════════════════════════════════════
//  SEPARATOR ESCAPING  (for TEXTJOIN)
// ══════════════════════════════════════════════════════════

/**
 * Escape a separator string so it can be safely embedded inside an Excel
 * double-quoted string argument.  e.g.  "  →  ""   \  stays as-is
 */
function escapeSeparator(sep) {
  return sep.replace(/"/g, '""');
}

// ══════════════════════════════════════════════════════════
//  COLUMN INDEX PARSER  (VLOOKUP)
// ══════════════════════════════════════════════════════════

/**
 * Parse comma-separated column indices.
 * Returns { values: number[], valueExpr: string, isMultiple: boolean, error? }
 */
function parseColumnIndex(input) {
  if (!input || !input.trim()) return { error: 'Column index is required.' };
  const nums = input.replace(/\s/g, '').split(',')
    .map(v => parseInt(v, 10))
    .filter(n => !isNaN(n) && n > 0);
  if (!nums.length) return { error: 'Enter valid positive numbers.' };
  const unique = [...new Set(nums)];
  if (unique.length !== nums.length) return { error: 'Duplicate column indices found.' };
  const isMultiple = unique.length > 1;
  const valueExpr  = isMultiple ? `{${unique.join(',')}}` : String(unique[0]);
  return { values: unique, valueExpr, isMultiple };
}

/**
 * Validate column index against the actual table column count.
 * Returns a warning string or null.
 */
function warnColIndexOutOfRange(colIndexResult, rangeComp) {
  if (!colIndexResult.values) return null;
  // Estimate table width if we have both start and end columns
  const colToNum = c => c.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const width = colToNum(rangeComp.endCol) - colToNum(rangeComp.startCol) + 1;
  if (width < 2) return null; // can't reliably tell
  const maxIdx = Math.max(...colIndexResult.values);
  if (maxIdx > width) {
    return `⚠️ Column index ${maxIdx} exceeds the estimated table width (${width} columns from ${rangeComp.startCol} to ${rangeComp.endCol}). Verify your range and index.`;
  }
  return null;
}

// ══════════════════════════════════════════════════════════
//  LAST ROW DETECTION
// ══════════════════════════════════════════════════════════

/**
 * Read the last-row radio group for a card prefix and return resolution info.
 * @param {string} prefix   — card prefix e.g. "tj"
 * @param {string} refCol   — column letter to use in MATCH/COUNTA
 * @param {string} sheetPath — full sheet path prefix (may be "")
 * @returns {{ ok, mode:'literal'|'formula', formula:string, error? }}
 */
function getLastRowInfo(prefix, refCol, sheetPath) {
  const mode = document.querySelector(`input[name="${prefix}-lastrow"]:checked`)?.value || 'match';

  if (mode === 'manual') {
    const raw = document.getElementById(`${prefix}-manualrow`)?.value?.trim();
    if (!raw || isNaN(raw) || parseInt(raw) < 2) {
      return { ok: false, error: 'Manual end row: enter a number ≥ 2.' };
    }
    return { ok: true, mode: 'literal', formula: raw };
  }

  if (mode === 'counta') {
    return { ok: true, mode: 'formula',
      formula: `COUNTA(${sheetPath}$${refCol}:$${refCol})` };
  }

  // MATCH — robust, handles text + numbers
  return { ok: true, mode: 'formula',
    formula: `MATCH(2,1/(${sheetPath}$${refCol}:$${refCol}<>""))` };
}

// ══════════════════════════════════════════════════════════
//  RANGE BUILDER HELPERS
// ══════════════════════════════════════════════════════════

/**
 * Build an INDIRECT-wrapped dynamic range string.
 * e.g. INDIRECT("'Sheet1'!$B$2:$F"&lastRow)
 */
function buildIndirectRange(comp, defStart, lrExpr, sheetPath, lock) {
  const lc = lock ? '$' : '';
  return `INDIRECT("${sheetPath}${lc}${comp.startCol}${lc}${defStart}:${lc}${comp.endCol}"&${lrExpr})`;
}

/**
 * Build a plain static range string (no INDIRECT).
 * e.g.  'Sheet1'!$A$2:$D$200
 */
function buildStaticRange(comp, defStart, endRow, sheetPath, lock) {
  const lc = lock ? '$' : '';
  return `${sheetPath}${lc}${comp.startCol}${lc}${defStart}:${lc}${comp.endCol}${lc}${endRow}`;
}

// ══════════════════════════════════════════════════════════
//  ERROR WRAPPING
// ══════════════════════════════════════════════════════════

/**
 * Wrap a formula fragment with IFNA / IFERROR.
 * Order: IFERROR outermost, IFNA next.
 */
function wrapErrors(inner, ifnaVal, ierrVal) {
  let f = inner;
  if (ifnaVal && ifnaVal.trim()) f = `IFNA(${f},"${ifnaVal.trim().replace(/"/g,'""')}")`;
  if (ierrVal && ierrVal.trim()) f = `IFERROR(${f},"${ierrVal.trim().replace(/"/g,'""')}")`;
  return f;
}

// ══════════════════════════════════════════════════════════
//  UI HELPERS
// ══════════════════════════════════════════════════════════

function browseFile(inputId) {
  const inp = document.createElement('input');
  inp.type = 'file';
  inp.accept = '.xlsx,.xls,.xlsm,.xlsb';
  inp.onchange = e => {
    if (e.target.files[0]) document.getElementById(inputId).value = e.target.files[0].name;
  };
  inp.click();
}

function copyFormula(textareaId, btn) {
  const val = document.getElementById(textareaId)?.value;
  if (!val) return;
  navigator.clipboard.writeText(val).then(() => {
    const orig = btn.textContent;
    btn.textContent = '✅ Copied!';
    btn.classList.add('copied');
    setTimeout(() => { btn.textContent = orig; btn.classList.remove('copied'); }, 2000);
  }).catch(() => {
    const ta = document.getElementById(textareaId);
    if (ta) { ta.select(); document.execCommand('copy'); }
  });
}

function copyHelper(codeId) {
  const code = document.getElementById(codeId);
  if (!code) return;
  navigator.clipboard.writeText(code.textContent.trim()).then(() => {
    const btn = event.currentTarget;
    const orig = btn.textContent;
    btn.textContent = '✅ Copied!';
    setTimeout(() => { btn.textContent = orig; }, 2000);
  });
}

/** Switch output tabs within a card */
function switchTab(prefix, tabName, allTabs) {
  allTabs.forEach(t => {
    document.getElementById(`${prefix}-tab-${t}`)?.classList.remove('active');
    document.getElementById(`${prefix}-pane-${t}`)?.classList.remove('active');
  });
  document.getElementById(`${prefix}-tab-${tabName}`)?.classList.add('active');
  document.getElementById(`${prefix}-pane-${tabName}`)?.classList.add('active');
}

/** Show/hide manual row input when radio changes */
function onLastRowChange(prefix) {
  const val  = document.querySelector(`input[name="${prefix}-lastrow"]:checked`)?.value;
  const wrap = document.getElementById(`${prefix}-manualrow-wrap`);
  if (wrap) wrap.classList.toggle('show', val === 'manual');
}

/**
 * Show an inline warning banner inside a pane.
 * @param {string} elId — id of a container element to prepend warning into
 * @param {string} msg  — warning message
 */
function showPaneWarning(elId, msg) {
  const el = document.getElementById(elId);
  if (!el || !msg) return;
  // Remove previous warnings
  el.querySelectorAll('.inline-warn').forEach(w => w.remove());
  const div = document.createElement('div');
  div.className = 'inline-warn';
  div.textContent = msg;
  el.prepend(div);
}

function clearPaneWarning(elId) {
  document.getElementById(elId)?.querySelectorAll('.inline-warn').forEach(w => w.remove());
}

// ══════════════════════════════════════════════════════════
//  RESET
// ══════════════════════════════════════════════════════════

function resetFormByPrefix(prefix, onAfterReset) {
  document.querySelectorAll(`[id^="${prefix}-"]`).forEach(el => {
    if      (el.tagName === 'SELECT') el.selectedIndex = 0;
    else if (el.type === 'checkbox')  el.checked = (el.dataset.default !== 'false');
    else if (el.type === 'radio')     { /* handled below */ }
    else if (el.type === 'number')    el.value = '';
    else if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') el.value = '';
  });
  // Reset last-row radio to first option
  const radios = document.querySelectorAll(`input[name="${prefix}-lastrow"]`);
  radios.forEach(r => r.checked = false);
  if (radios[0]) { radios[0].checked = true; onLastRowChange(prefix); }
  // Hide output section
  document.getElementById(`${prefix}-output-section`)?.classList.remove('visible');
  if (typeof onAfterReset === 'function') onAfterReset();
}

// ══════════════════════════════════════════════════════════
//  CARD COLLAPSE / EXPAND
// ══════════════════════════════════════════════════════════

function toggleCard(headerEl) {
  headerEl.closest('.formula-card').classList.toggle('collapsed');
  saveCardStates();
}
function expandAllCards() {
  document.querySelectorAll('.formula-card').forEach(c => c.classList.remove('collapsed'));
  saveCardStates();
}
function collapseAllCards() {
  document.querySelectorAll('.formula-card').forEach(c => c.classList.add('collapsed'));
  saveCardStates();
}
function saveCardStates() {
  try {
    const states = [...document.querySelectorAll('.formula-card[data-card-id]')].map(c => ({
      id: c.dataset.cardId, collapsed: c.classList.contains('collapsed')
    }));
    localStorage.setItem('ahmos_cardStates', JSON.stringify(states));
  } catch(e) {}
}
function loadCardStates() {
  try {
    const saved = JSON.parse(localStorage.getItem('ahmos_cardStates') || '[]');
    saved.forEach(s => {
      const el = document.querySelector(`.formula-card[data-card-id="${s.id}"]`);
      if (el && s.collapsed) el.classList.add('collapsed');
    });
  } catch(e) {}
}

// ══════════════════════════════════════════════════════════
//  DRAG-TO-REORDER CARDS
// ══════════════════════════════════════════════════════════

let _dragSrc = null;

function initCardDrag() {
  const container = document.getElementById('cards-container');
  if (!container) return;

  // Load saved order first
  try {
    const savedOrder = JSON.parse(localStorage.getItem('ahmos_cardOrder') || '[]');
    if (savedOrder.length) {
      savedOrder.forEach(id => {
        const card = container.querySelector(`.formula-card[data-card-id="${id}"]`);
        if (card) container.appendChild(card);
      });
    }
  } catch(e) {}

  container.addEventListener('dragstart', e => {
    const card = e.target.closest('.formula-card');
    if (!card) return;
    _dragSrc = card;
    card.classList.add('dragging');
    e.dataTransfer.effectAllowed = 'move';
  });

  container.addEventListener('dragend', e => {
    const card = e.target.closest('.formula-card');
    if (card) card.classList.remove('dragging');
    container.querySelectorAll('.formula-card').forEach(c => c.classList.remove('drag-over'));
    saveCardOrder();
  });

  container.addEventListener('dragover', e => {
    e.preventDefault();
    const card = e.target.closest('.formula-card');
    if (!card || card === _dragSrc) return;
    container.querySelectorAll('.formula-card').forEach(c => c.classList.remove('drag-over'));
    card.classList.add('drag-over');
    // Determine insert position
    const rect = card.getBoundingClientRect();
    const mid  = rect.top + rect.height / 2;
    if (e.clientY < mid) {
      container.insertBefore(_dragSrc, card);
    } else {
      container.insertBefore(_dragSrc, card.nextSibling);
    }
  });

  container.addEventListener('dragleave', e => {
    const card = e.target.closest('.formula-card');
    if (card) card.classList.remove('drag-over');
  });

  container.addEventListener('drop', e => { e.preventDefault(); });
}

function saveCardOrder() {
  try {
    const order = [...document.querySelectorAll('#cards-container .formula-card[data-card-id]')]
      .map(c => c.dataset.cardId);
    localStorage.setItem('ahmos_cardOrder', JSON.stringify(order));
  } catch(e) {}
}

// ══════════════════════════════════════════════════════════
//  SHARED — LIVE PREVIEW ENGINE
// ══════════════════════════════════════════════════════════

const _previewTimers = {};

/**
 * Schedule a debounced preview for a card.
 * @param {string} prefix   — card prefix e.g. "tj"
 * @param {Function} buildFn — () => string | null  (pure formula builder, no DOM side-effects)
 */
function schedulePreview(prefix, buildFn) {
  clearTimeout(_previewTimers[prefix]);
  _previewTimers[prefix] = setTimeout(() => runPreview(prefix, buildFn), 280);
}

function runPreview(prefix, buildFn) {
  const el = document.getElementById(`${prefix}-live-preview`);
  if (!el) return;
  const hint = msg => { el.textContent = msg; el.className = 'live-preview-box hint'; };
  try {
    const result = buildFn();
    if (!result) return hint('— fill in required fields to see a preview —');
    // Show first formula line only (strip comment lines)
    const firstLine = result.split('\n').find(l => l.startsWith('=')) || result;
    el.textContent = firstLine.length > 200 ? firstLine.slice(0, 200) + '…' : firstLine;
    el.className = 'live-preview-box ready';
  } catch(e) {
    hint('— fill in required fields to see a preview —');
  }
}

// ══════════════════════════════════════════════════════════
//  SHARED — FORMULA EXPLANATION ENGINE
// ══════════════════════════════════════════════════════════

/**
 * Function reference database used by all cards.
 * Each card passes the list of function names it uses.
 */
const FUNC_DB = {
  LET     : { color:'#c084fc', what:'LET(name,value,…,formula) — assigns names to intermediate values within the formula.', why:'Avoids repeating the same sub-expression multiple times, making the formula shorter, faster to calculate, and easier to read.' },
  INDIRECT: { color:'#38c8ff', what:'INDIRECT(text) — converts a text string into a live cell or range reference.', why:'Lets us build a dynamic range whose end row changes at runtime (e.g. "A2:A"&lastRow). Without INDIRECT the range would be fixed at formula entry time.' },
  FILTER  : { color:'#4fffb0', what:'FILTER(array, include, [if_empty]) — returns only the rows where the condition is TRUE.', why:'The core filter engine. Replaces the need for helper columns or complex array formulas to pick matching rows.' },
  TEXTJOIN: { color:'#ffd166', what:'TEXTJOIN(delimiter, ignore_empty, text…) — joins multiple values into one string with a separator.', why:'Collapses the array of matching values returned by FILTER into a single readable cell instead of spilling across multiple rows.' },
  XLOOKUP : { color:'#f472b6', what:'XLOOKUP(lookup, lookup_array, return_array, [not_found], [match_mode], [search_mode]) — finds a value and returns a result from another range.', why:'The modern replacement for VLOOKUP. Supports reverse search, wildcard matching, and returns from any column — not just to the right.' },
  VLOOKUP : { color:'#fb923c', what:'VLOOKUP(lookup, table, col_index, [match_type]) — searches the first column of a table and returns a value from a specified column.', why:'Classic lookup function. Works in all Excel versions. The column to return must be to the right of the lookup column.' },
  INDEX   : { color:'#a78bfa', what:'INDEX(array, row_num) — returns the value at a given position in a range.', why:'The return half of the INDEX+MATCH pair. Unlike VLOOKUP it can return from any column regardless of position.' },
  MATCH   : { color:'#38c8ff', what:'MATCH(lookup, array, [match_type]) — returns the relative position of a value in a range.', why:'The locate half of the INDEX+MATCH pair. Finds the row number that INDEX then uses to retrieve the actual value.' },
  IFERROR : { color:'#60a5fa', what:'IFERROR(value, fallback) — catches any formula error and returns the fallback instead.', why:'Primary error trap. Covers #N/A, #VALUE!, #REF!, #NAME? and all other error types.' },
  IFNA    : { color:'#818cf8', what:'IFNA(value, fallback) — catches only #N/A errors and returns the fallback.', why:'Narrower than IFERROR — used specifically to signal "not found" without masking other unexpected errors.' },
  SUMIF   : { color:'#4fffb0', what:'SUMIF(range, criteria, sum_range) — sums cells where the criteria matches.', why:'Single-condition sum. Faster and simpler than SUMIFS when only one filter is needed.' },
  SUMIFS  : { color:'#4fffb0', what:'SUMIFS(sum_range, criteria_range1, criteria1, …) — sums cells that meet multiple conditions.', why:'Multi-condition version of SUMIF. All conditions must be true simultaneously (AND logic).' },
  COUNTIF : { color:'#fbbf24', what:'COUNTIF(range, criteria) — counts cells that match a single condition.', why:'Single-condition count. Returns 0 when nothing matches, making error wrapping less critical.' },
  COUNTIFS: { color:'#fbbf24', what:'COUNTIFS(range1, criteria1, range2, criteria2, …) — counts cells meeting multiple conditions.', why:'AND-logic count across multiple criteria ranges. Returns 0 if no rows satisfy all conditions.' },
  TRIM    : { color:'#4fffb0', what:'TRIM(text) — removes leading, trailing, and extra internal spaces.', why:'Extracted values from MID can include unwanted spaces depending on source data formatting.' },
  MID     : { color:'#ffd166', what:'MID(text, start, num_chars) — extracts a substring starting at a given position.', why:'Core extractor. Start position and length are calculated dynamically via FIND.' },
  FIND    : { color:'#fb923c', what:'FIND(find_text, within_text, [start]) — returns the position of a string inside text. Case-sensitive.', why:'Used as the anchor: locates the search term and the separator so MID knows where to start and stop.' },
  VALUE   : { color:'#94a3b8', what:'VALUE(text) — converts a number stored as text into a numeric value.', why:'Used in the 4-cast VLOOKUP chain to try numeric lookup when text-cast versions still fail.' },
  ROW     : { color:'#94a3b8', what:'ROW([reference]) — returns the row number of a cell or range.', why:'Used in the legacy IF+SMALL+INDEX pattern to convert absolute row numbers to relative positions within the data range.' },
  SMALL   : { color:'#94a3b8', what:'SMALL(array, k) — returns the k-th smallest value in a set.', why:'Part of the legacy multi-match trick: extracts the 1st, 2nd, 3rd… matching row index when copied down.' },
  IF      : { color:'#94a3b8', what:'IF(condition, value_if_true, value_if_false) — conditional branching.', why:'Used in the legacy array formula to build a list of row positions where the condition is met.' },
};

/**
 * Build the explanation panel HTML for a card.
 * @param {string[]} funcNames  — ordered list of function names to explain
 * @param {string}   introText  — one sentence describing what the formula does
 * @param {string[]} [notes]    — optional extra note paragraphs
 */
function buildExplainHTML(funcNames, introText, notes = []) {
  const chain = funcNames.join(' → ');
  let html = `
    <div class="explain-intro">
      ${introText}<br>
      <span style="opacity:.6;font-size:11px">Structure:</span>
      <code class="explain-code">${chain}</code>
    </div>
    <div class="explain-table">`;

  funcNames.forEach(fn => {
    const info = FUNC_DB[fn];
    if (!info) return;
    html += `
      <div class="explain-row">
        <span class="explain-fn-badge" style="background:${info.color}18;border-color:${info.color}60;color:${info.color}">${fn}</span>
        <div class="explain-detail">
          <div class="explain-what"><strong>What:</strong> ${info.what}</div>
          <div class="explain-why"><strong>Why here:</strong> ${info.why}</div>
        </div>
      </div>`;
  });

  html += `</div>`;
  notes.forEach(n => { html += `<div class="explain-note">${n}</div>`; });
  return html;
}

/**
 * Toggle explanation panel open/closed.
 * @param {string}   prefix
 * @param {Function} buildFn — () => string (returns the HTML for the panel)
 */
function toggleExplanation(prefix, buildFn) {
  const panel = document.getElementById(`${prefix}-explain-panel`);
  const btn   = document.getElementById(`${prefix}-explain-btn`);
  if (!panel) return;
  const open = panel.classList.contains('open');
  if (open) {
    panel.classList.remove('open');
    setTimeout(() => { if (!panel.classList.contains('open')) panel.innerHTML = ''; }, 320);
    if (btn) btn.textContent = '💡 Explain Formula';
  } else {
    panel.innerHTML = buildFn();
    panel.offsetHeight; // force reflow for CSS transition
    panel.classList.add('open');
    if (btn) btn.textContent = '✕ Hide Explanation';
  }
}

// ══════════════════════════════════════════════════════════
//  SHARED — EXPORT ALL
// ══════════════════════════════════════════════════════════

/**
 * Collect all textarea values from an output section and export as clipboard/txt.
 * @param {string}   prefix    — card prefix
 * @param {string}   cardLabel — human-readable card name for the header
 * @param {Array}    tabs      — [{id, label}] — id is the textarea id, label is the output name
 */
function exportAll(prefix, cardLabel, tabs) {
  const stamp = new Date().toLocaleString();
  const lines = [
    '══════════════════════════════════════════════════════',
    `  Ahmos Excel Formula Generator  —  ${cardLabel}`,
    `  Exported: ${stamp}`,
    '══════════════════════════════════════════════════════',
  ];

  let hasContent = false;
  tabs.forEach(({ id, label }) => {
    const el  = document.getElementById(id);
    const val = el?.value?.trim();
    if (!val) return;
    hasContent = true;
    lines.push('', `── ${label} ──`, val);
  });

  if (!hasContent) return;
  lines.push('', '══════════════════════════════════════════════════════');
  const full = lines.join('\n');

  const btn   = document.getElementById(`${prefix}-export-btn`);
  const flash = msg => {
    if (!btn) return;
    const orig = btn.textContent;
    btn.textContent = msg;
    setTimeout(() => { btn.textContent = orig; }, 2500);
  };

  navigator.clipboard.writeText(full)
    .then(() => flash('✅ Copied!'))
    .catch(() => {
      const blob = new Blob([full], { type: 'text/plain;charset=utf-8' });
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
      a.href = url; a.download = `ahmos-${prefix}-formulas.txt`; a.click();
      URL.revokeObjectURL(url);
      flash('✅ Downloaded!');
    });
}

// ══════════════════════════════════════════════════════════
//  RESET ALL STORED SETTINGS
// ══════════════════════════════════════════════════════════

/**
 * Clear every ahmos_* localStorage key (card order, collapse states, presets).
 * Prompts the user first, then reloads so defaults are restored cleanly.
 */
function resetAllSettings() {
  const msg = [
    'This will reset:',
    '  • Card arrangement order',
    '  • Collapsed / expanded card states',
    '  • Extract Text custom presets',
    '',
    'All form inputs stay untouched.',
    '',
    'Continue?'
  ].join('\n');
  if (!confirm(msg)) return;
  try {
    Object.keys(localStorage)
      .filter(k => k.startsWith('ahmos_'))
      .forEach(k => localStorage.removeItem(k));
  } catch(e) {}
  location.reload();
}

// ══════════════════════════════════════════════════════════
//  BOOT
// ══════════════════════════════════════════════════════════

document.addEventListener('DOMContentLoaded', () => {
  loadCardStates();
  initCardDrag();
});
