/* ═══════════════════════════════════════════════════════════
   card-vl.js — VLOOKUP  (V6.2 — full review pass)

   "Force text coercion" ON means BOTH SIDES:
     SEARCH side  → lookup value coerced:  A2&""  (text match)
                                           A2*1   (numeric match)
     RESULT side  → returned value coerced: VLOOKUP(...)&""
                   so the output is always a clean text string

   Outputs:
     Advanced  — 4-cast chain (raw/text/×1/VALUE) + result &""
     Standard  — 2-cast (text + numeric) + result &""  [simpler, same safety]
     Legacy    — single text-cast + result &"", full-column table

   addStr OFF → single clean VLOOKUP, no coercion anywhere.
   ═══════════════════════════════════════════════════════════ */
'use strict';

const VL_TABS = ['adv', 'c19', 'leg'];
function vl_switchTab(tab) { switchTab('vl', tab, VL_TABS); }

function vl_reset() {
  resetFormByPrefix('vl', () => {
    document.getElementById('vl-match').selectedIndex = 0;
    document.getElementById('vl-addstring').checked   = true;
    document.getElementById('vl-lock').checked        = true;
    document.getElementById('vl-usetrim').checked     = false;
  });
  const pr = document.getElementById('vl-live-preview');
  if (pr) { pr.textContent = '— fill in the fields above to see a live preview —'; pr.className = 'live-preview-box hint'; }
  const panel = document.getElementById('vl-explain-panel');
  if (panel) { panel.classList.remove('open'); panel.innerHTML = ''; }
  const btn = document.getElementById('vl-explain-btn');
  if (btn) btn.textContent = '💡 Explain Formula';
}

function vl_getInputs() {
  return {
    filePath  : document.getElementById('vl-filepath').value.trim(),
    sheetName : document.getElementById('vl-sheetname').value.trim(),
    lookupRaw : document.getElementById('vl-lookupcell').value.trim(),
    rangeRaw  : document.getElementById('vl-range').value.trim(),
    colIdxRaw : document.getElementById('vl-colindex').value.trim(),
    matchType : document.getElementById('vl-match').value,
    ifnaVal   : document.getElementById('vl-ifna').value.trim(),
    ierrVal   : document.getElementById('vl-error').value.trim(),
    lock      : document.getElementById('vl-lock').checked,
    addStr    : document.getElementById('vl-addstring').checked,
    useTrim   : document.getElementById('vl-usetrim').checked,
  };
}

function vl_buildAll() {
  const { filePath, sheetName, lookupRaw, rangeRaw, colIdxRaw,
          matchType, ifnaVal, ierrVal, lock, addStr, useTrim } = vl_getInputs();

  if (!lookupRaw || !rangeRaw || !colIdxRaw) return null;

  const lookupCell = normaliseCellRef(lookupRaw);
  if (!/^[A-Z]{1,3}\d{1,7}$/.test(lookupCell)) return null;

  const rangeV = validateRange(rangeRaw, 'Table Range');
  if (!rangeV.ok) return null;

  const rc       = parseRange(rangeRaw);
  const ciResult = parseColumnIndex(colIdxRaw);
  if (ciResult.error) return null;

  const sheetPath = buildSheetPath(filePath, sheetName);
  if (sheetPath === null) return null;

  const lc       = lock ? '$' : '';
  const defStart = rc.startRow || '1';

  // Escape error/ifna text for embedding in Excel double-quoted strings
  const errQ = (ierrVal || 'Not Found').replace(/"/g, '""');
  const naQ  = (ifnaVal || '').replace(/"/g, '""');

  const idx    = ciResult.valueExpr;
  const isMulti = ciResult.isMultiple;
  const legIdx  = isMulti ? ciResult.values[0] : idx;

  // ── Table references ──────────────────────────────────────
  // Dynamic (INDIRECT) for Adv/Standard when no explicit end row;
  // full-column static for Legacy (always safe, no INDIRECT needed).
  let tableRef, tableRefFull;
  if (rc.endRow) {
    tableRef = tableRefFull =
      `${sheetPath}${lc}${rc.startCol}${lc}${defStart}:${lc}${rc.endCol}${lc}${rc.endRow}`;
  } else {
    const lr = getLastRowInfo('vl', rc.endCol, sheetPath);
    if (!lr.ok) return null;
    tableRef     = buildIndirectRange(rc, defStart, lr.formula, sheetPath, lock);
    tableRefFull = `${sheetPath}${lc}${rc.startCol}:${lc}${rc.endCol}`;
  }

  // ── Lookup value base (optional TRIM) ─────────────────────
  const lv_base = useTrim ? `TRIM(${lookupCell})` : lookupCell;

  // ── VLOOKUP shorthand ─────────────────────────────────────
  const VL = (lv, tbl, i) => `VLOOKUP(${lv},${tbl},${i},${matchType})`;

  // ── Outer error-wrap helper ───────────────────────────────
  // IFNA is always inner (catches #N/A from VLOOKUP — only when user set ifnaVal).
  // IFERROR is always outermost (catches everything else — always present).
  // errQ is the user's iferror value, or "Not Found" as a guaranteed fallback.
  const wrap = (inner) => {
    let f = inner;
    if (ifnaVal) f = `IFNA(${f},"${naQ}")`;
    f = `IFERROR(${f},"${errQ}")`; // always wrap — errQ already has a default
    return f;
  };

  let adv, advDesc, c19, c19Desc, leg, legDesc;

  if (addStr) {
    // ╔══ ADVANCED — 4-cast chain, result coerced to text ══════════╗
    // Tries every type combination in sequence. Each IFERROR falls
    // through to the next. The entire chain is then coerced with &""
    // so the RETURNED VALUE is also always text.
    //
    // =IFERROR( (IFERROR(VL(raw), IFERROR(VL(txt), IFERROR(VL(×1), VL(VALUE)))))&"", "err" )
    const inner4 =
      `IFERROR(${VL(lv_base,          tableRef, idx)},` +
      `IFERROR(${VL(`${lv_base}&""`,  tableRef, idx)},` +
      `IFERROR(${VL(`${lv_base}*1`,   tableRef, idx)},` +
             `${VL(`VALUE(${lv_base})`, tableRef, idx)})))`;
    // Coerce result to text (RESULT side), then wrap errors
    adv     = `=${wrap(`(${inner4})&""`)}`;
    advDesc = 'Force text ON — both sides: 4-cast chain on lookup (raw → text → ×1 → VALUE), result coerced to text with &"". Catches every type-mismatch silently. Dynamic range via INDIRECT.';

    // ╔══ STANDARD — 2-cast, result coerced ════════════════════════╗
    // Simpler: try text cast first, fall back to numeric cast.
    // Result is also coerced to text.
    const inner2 =
      `IFERROR(${VL(`${lv_base}&""`, tableRef, idx)},` +
              `${VL(`${lv_base}*1`,  tableRef, idx)})`;
    c19     = `=${wrap(`(${inner2})&""`)}`;
    c19Desc = 'Force text ON — both sides: text-cast (&"") + numeric fallback (×1) on lookup, result coerced to text (&""). Handles mixed number/text data. Dynamic range via INDIRECT.';

    // ╔══ LEGACY — text-cast + result coerced, full-column table ═══╗
    leg     = `=${wrap(`${VL(`${lv_base}&""`, tableRefFull, legIdx)}&""`)}`;
    legDesc = 'Force text ON — both sides: lookup coerced to text (&""), result coerced to text (&""). Full-column reference. Excel 2003+.';

  } else {
    // ╔══ No coercion — single clean VLOOKUP ═══════════════════════╗
    adv     = `=${wrap(VL(lv_base, tableRef, idx))}`;
    advDesc = 'Force text OFF: single clean VLOOKUP, dynamic range via INDIRECT. Ensure lookup value and table first column share the same data type.';

    c19     = `=${wrap(VL(lv_base, tableRef, idx))}`;
    c19Desc = 'Force text OFF: plain VLOOKUP with INDIRECT dynamic range. No coercion applied.';

    leg     = `=${wrap(VL(lv_base, tableRefFull, legIdx))}`;
    legDesc = 'Force text OFF: plain VLOOKUP, full-column reference. Excel 2003+.';
  }

  const rangeWarn = warnColIndexOutOfRange(ciResult, rc);

  return {
    adv, advDesc, c19, c19Desc, leg, legDesc,
    rangeWarn,
    isMultiIdx : isMulti,
    firstIdx   : ciResult.values?.[0],
    allIdx     : ciResult.values?.join(','),
  };
}

// ── Live preview ───────────────────────────────────────────
function vl_previewBuildFn() {
  try { const r = vl_buildAll(); return r?.adv || null; } catch(e) { return null; }
}

// ── Explanation ────────────────────────────────────────────
function vl_explainBuildFn() {
  const addStr  = document.getElementById('vl-addstring')?.checked;
  const useTrim = document.getElementById('vl-usetrim')?.checked;
  const funcs   = addStr
    ? ['IFERROR','IFNA','VALUE','VLOOKUP','INDIRECT']
    : ['IFERROR','IFNA','VLOOKUP','INDIRECT'];
  const notes = [
    addStr
      ? '📝 <strong>Force text ON — both sides:</strong>' +
        ' Lookup value is tried as text (<code>&amp;""</code>) and as a number (<code>×1</code>) to match regardless of how Excel stored the data.' +
        ' The returned value is also cast to text with <code>&amp;""</code> — so the output is always a clean string on both ends.'
      : '📝 <strong>Force text OFF:</strong> Single clean VLOOKUP, no coercion.' +
        ' Ensure your lookup value and the first column of your table store data in the same type (both text or both numbers).',
    useTrim
      ? '✂️ <strong>TRIM ON:</strong> Whitespace stripped from the lookup value before searching — prevents invisible space mismatches.'
      : '',
  ].filter(Boolean);
  return buildExplainHTML(funcs,
    'Searches the <strong>first column</strong> of a table for a lookup value, then returns the value from a specified column in the same row.',
    notes);
}

// ── Init ───────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  ['vl-filepath','vl-sheetname','vl-lookupcell','vl-range',
   'vl-colindex','vl-ifna','vl-error','vl-manualrow']
    .forEach(id => document.getElementById(id)
      ?.addEventListener('input', () => schedulePreview('vl', vl_previewBuildFn)));
  ['vl-match','vl-addstring','vl-lock','vl-usetrim']
    .forEach(id => document.getElementById(id)
      ?.addEventListener('change', () => schedulePreview('vl', vl_previewBuildFn)));
  document.querySelectorAll('input[name="vl-lastrow"]')
    .forEach(r => r.addEventListener('change',
      () => schedulePreview('vl', vl_previewBuildFn)));
});

// ── Generator ──────────────────────────────────────────────
function generateVLookup() {
  const { lookupRaw, rangeRaw, colIdxRaw, filePath, sheetName } = vl_getInputs();

  if (!lookupRaw || !rangeRaw || !colIdxRaw) {
    alert('Please fill in: Lookup Cell, Table Range, and Column Index.'); return;
  }
  const lookupNorm = normaliseCellRef(lookupRaw);
  if (!/^[A-Z]{1,3}\d{1,7}$/.test(lookupNorm)) {
    alert(`Lookup Cell "${lookupRaw}" is not a valid cell reference.\nExample: A2`); return;
  }
  const rangeV = validateRange(rangeRaw, 'Table Range');
  if (!rangeV.ok) { alert(rangeV.error); return; }

  const ci = parseColumnIndex(colIdxRaw);
  if (ci.error) { alert(`Column Index: ${ci.error}`); return; }

  const sp = buildSheetPath(filePath, sheetName);
  if (sp === null) { alert(processFilePath(filePath).error); return; }

  const r = vl_buildAll();
  if (!r) { alert('Could not build formula — please check your inputs.'); return; }

  document.getElementById('vl-out-adv').value        = r.adv;
  document.getElementById('vl-desc-adv').textContent = r.advDesc;
  document.getElementById('vl-out-c19').value        = r.c19;
  document.getElementById('vl-desc-c19').textContent = r.c19Desc;
  document.getElementById('vl-out-leg').value        = r.leg;
  document.getElementById('vl-desc-leg').textContent = r.legDesc;

  ['adv','c19','leg'].forEach(t => {
    const warns = [];
    if (r.rangeWarn) warns.push(r.rangeWarn);
    if (t === 'leg' && r.isMultiIdx)
      warns.push(
        `⚠️ Multiple column indices {${r.allIdx}} require dynamic arrays — not available in legacy.` +
        ` Only the first index (${r.firstIdx}) is used here. Use Advanced or Standard for multi-column returns.`
      );
    warns.length
      ? showPaneWarning(`vl-pane-${t}`, warns.join('\n\n'))
      : clearPaneWarning(`vl-pane-${t}`);
  });

  document.getElementById('vl-output-section').classList.add('visible');
  vl_switchTab('adv');

  const panel = document.getElementById('vl-explain-panel');
  if (panel?.classList.contains('open')) panel.innerHTML = vl_explainBuildFn();

  document.getElementById('vl-output-section')
    .scrollIntoView({ behavior:'smooth', block:'nearest' });
}
