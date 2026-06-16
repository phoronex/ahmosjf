/* ═══════════════════════════════════════════════════════════
   card-tj.js — TEXTJOIN + FILTER
   • Force text coercion: ""& on BOTH match value AND range arrays
   • TRIM option on returned values
   • Live preview  •  Explain Formula  •  Export All
   ═══════════════════════════════════════════════════════════ */
'use strict';

const TJ_TABS = ['adv', 'c19', 'leg'];
function tj_switchTab(tab) { switchTab('tj', tab, TJ_TABS); }

function tj_reset() {
  resetFormByPrefix('tj', () => {
    document.getElementById('tj-ignoreblank').value   = 'TRUE';
    document.getElementById('tj-usetextjoin').checked = true;
    document.getElementById('tj-addstring').checked   = true;
    document.getElementById('tj-lock').checked        = true;
    document.getElementById('tj-usetrim').checked     = false;
  });
  const pr = document.getElementById('tj-live-preview');
  if (pr) { pr.textContent = '— fill in the fields above to see a live preview —'; pr.className = 'live-preview-box hint'; }
  const panel = document.getElementById('tj-explain-panel');
  if (panel) { panel.classList.remove('open'); panel.innerHTML = ''; }
  const btn = document.getElementById('tj-explain-btn');
  if (btn) btn.textContent = '💡 Explain Formula';
}

// ── Pure builder — returns {adv,c19,legA,legB} or throws ──
function tj_buildAll() {
  const filePath    = document.getElementById('tj-filepath').value.trim();
  const sheetName   = document.getElementById('tj-sheetname').value.trim();
  const returnRaw   = document.getElementById('tj-returncol').value.trim();
  const targetRaw   = document.getElementById('tj-targetcol').value.trim();
  const equalRaw    = document.getElementById('tj-equalcell').value.trim();
  const sepRaw      = document.getElementById('tj-separator').value || '|';
  const ifnaVal     = document.getElementById('tj-ifna').value.trim();
  const ierrVal     = document.getElementById('tj-error').value.trim();
  const ignoreBlank = document.getElementById('tj-ignoreblank').value;
  const lock        = document.getElementById('tj-lock').checked;
  const useTJ       = document.getElementById('tj-usetextjoin').checked;
  const addStr      = document.getElementById('tj-addstring').checked;
  const useTrim     = document.getElementById('tj-usetrim').checked;

  if (!returnRaw || !targetRaw || !equalRaw) return null;
  const retV = validateRange(returnRaw, 'Return Column');  if (!retV.ok) return null;
  const tgtV = validateRange(targetRaw, 'Filter Column');  if (!tgtV.ok) return null;

  const ret = parseRange(returnRaw);
  const tgt = parseRange(targetRaw);
  if (tgt.startCol !== tgt.endCol) return null;

  const matchVal  = resolveMatchValue(equalRaw);
  const sep       = escapeSeparator(sepRaw);
  const sheetPath = buildSheetPath(filePath, sheetName);
  if (sheetPath === null) return null;

  const rStart   = parseInt(ret.startRow || '1');
  const tStart   = parseInt(tgt.startRow || '1');
  const defStart = String(Math.max(rStart, tStart));

  const explicitEnd = ret.endRow || tgt.endRow;
  let lastRow, needsVar = false;
  if (explicitEnd) {
    lastRow = explicitEnd;
  } else {
    const lr = getLastRowInfo('tj', tgt.endCol, sheetPath);
    if (!lr.ok) return null;
    lastRow  = lr.formula;
    needsVar = (lr.mode === 'formula');
  }

  const indirect = (comp, lrExpr) => buildIndirectRange(comp, defStart, lrExpr, sheetPath, lock);

  // Force text coercion: wrap BOTH the return range AND filter range with &""
  // so Excel treats them identically regardless of stored type.
  // Also wrap the match value's cell-ref side with &"" when it's a cell ref.
  const filterCore = (retName, tgtName) => {
    // Return array: optionally TRIM, optionally cast to text
    let retExpr = retName;
    if (useTrim)   retExpr = `TRIM(${retExpr})`;
    if (addStr)    retExpr = `${retExpr}&""`;

    // Filter condition: both sides as text when addStr is ON
    const eqL = matchVal.formula;
    let filterCond;
    if (addStr) {
      const lhText = `${tgtName}&""`;
      const rhText = matchVal.type === 'cell' ? `${eqL}&""` : eqL;
      filterCond = `${lhText}=${rhText}`;
    } else {
      filterCond = `${tgtName}=${eqL}`;
    }

    let f = `FILTER(${retExpr},${filterCond},"")`;
    if (useTJ) f = `TEXTJOIN("${sep}",${ignoreBlank},${f})`;
    const errText = ierrVal || 'Not Found'; // guaranteed fallback for wrapErrors
    return wrapErrors(f, ifnaVal, errText);
  };

  const lrExpr = needsVar ? 'lastRow' : lastRow;

  const adv = needsVar
    ? `=LET(lastRow,${lastRow},rng,${indirect(ret,'lastRow')},tgt,${indirect(tgt,'lastRow')},${filterCore('rng','tgt')})`
    : `=${filterCore(indirect(ret, lrExpr), indirect(tgt, lrExpr))}`;

  const c19 = `=${filterCore(indirect(ret, lastRow), indirect(tgt, lastRow))}`;

  const endRow   = explicitEnd || '10000';
  const retR     = buildStaticRange(ret, defStart, endRow, sheetPath, lock);
  const tgtR     = buildStaticRange(tgt, defStart, endRow, sheetPath, lock);
  const errStr   = ierrVal || 'Not Found';
  const eqF      = matchVal.formula;
  const lc       = lock ? '$' : '';
  const legA     = `=IFERROR(INDEX(${retR},MATCH(${eqF},${tgtR},0)),"${errStr}")`;
  const legB     = `=IFERROR(INDEX(${retR},SMALL(IF(${tgtR}=${eqF},ROW(${tgtR})-ROW(${sheetPath}${lc}${tgt.startCol}${lc}${defStart})+1),ROW(A1))),"")`;

  return { adv, c19, legA, legB, needsVar, endRow: explicitEnd || '10000', isFallback: !explicitEnd && needsVar };
}

function tj_previewBuildFn() {
  try { const r = tj_buildAll(); return r?.adv || null; } catch(e) { return null; }
}

function tj_explainBuildFn() {
  const useTJ  = document.getElementById('tj-usetextjoin')?.checked;
  const addStr = document.getElementById('tj-addstring')?.checked;
  const useTrim= document.getElementById('tj-usetrim')?.checked;
  const funcs  = ['LET','INDIRECT','FILTER', ...(useTJ ? ['TEXTJOIN'] : []), 'IFERROR','IFNA'];
  const notes  = [
    addStr
      ? '📝 <strong>Force text coercion ON:</strong> Both the return range and the filter range are cast to text with <code>&amp;""</code>, and the match value is also coerced. This ensures Excel treats stored numbers and text numbers identically.'
      : '📝 <strong>Force text coercion OFF:</strong> Ranges are compared as-is. Make sure your data types match to avoid missed matches.',
    useTrim
      ? '✂️ <strong>TRIM ON:</strong> Whitespace stripped from each returned value before joining.'
      : ''
  ].filter(Boolean);
  return buildExplainHTML(funcs,
    'Filters a range to rows where the filter column matches a value, then ' +
    (useTJ ? 'joins all matching return values into one cell with a separator.' : 'returns the matching return values as a spilled array.'),
    notes);
}


document.addEventListener('DOMContentLoaded', () => {
  const watch = ['tj-returncol','tj-targetcol','tj-equalcell','tj-separator',
                 'tj-ifna','tj-error','tj-filepath','tj-sheetname','tj-manualrow'];
  watch.forEach(id => document.getElementById(id)
    ?.addEventListener('input', () => schedulePreview('tj', tj_previewBuildFn)));
  ['tj-usetextjoin','tj-addstring','tj-lock','tj-usetrim'].forEach(id =>
    document.getElementById(id)
      ?.addEventListener('change', () => schedulePreview('tj', tj_previewBuildFn)));
  document.querySelectorAll('input[name="tj-lastrow"]')
    .forEach(r => r.addEventListener('change', () => schedulePreview('tj', tj_previewBuildFn)));
});

function generateTextJoin() {
  const filePath    = document.getElementById('tj-filepath').value.trim();
  const sheetName   = document.getElementById('tj-sheetname').value.trim();
  const returnRaw   = document.getElementById('tj-returncol').value.trim();
  const targetRaw   = document.getElementById('tj-targetcol').value.trim();
  const equalRaw    = document.getElementById('tj-equalcell').value.trim();
  const ifnaVal     = document.getElementById('tj-ifna').value.trim();
  const ierrVal     = document.getElementById('tj-error').value.trim();

  if (!returnRaw || !targetRaw || !equalRaw) {
    alert('Please fill in: Return Column, Filter Column, and Match Value / Cell.'); return;
  }
  const retV = validateRange(returnRaw, 'Return Column'); if (!retV.ok) { alert(retV.error); return; }
  const tgtV = validateRange(targetRaw, 'Filter Column'); if (!tgtV.ok) { alert(tgtV.error); return; }
  const tgt  = parseRange(targetRaw);
  if (tgt.startCol !== tgt.endCol) {
    alert(`Filter Column must be a single column.\nExample: A  or  A2:A`); return;
  }

  const sheetPath = buildSheetPath(filePath, sheetName);
  if (sheetPath === null) { alert(processFilePath(filePath).error); return; }

  const ret        = parseRange(returnRaw);
  const mismatchWarn = warnIfStartRowMismatch(ret, tgt);
  const lr_check   = getLastRowInfo('tj', tgt.endCol, sheetPath);
  if (!ret.endRow && !tgt.endRow && !lr_check.ok) { alert(lr_check.error); return; }

  const r = tj_buildAll();
  if (!r) { alert('Could not build formula — please check your inputs.'); return; }

  document.getElementById('tj-out-adv').value  = r.adv;
  document.getElementById('tj-desc-adv').textContent = 'LET + INDIRECT + FILTER' +
    (document.getElementById('tj-usetextjoin').checked?' + TEXTJOIN':'') +
    '. Dynamic range, named variables. Excel 365 / 2021+.';

  document.getElementById('tj-out-c19').value  = r.c19;
  document.getElementById('tj-desc-c19').textContent = 'FILTER' +
    (document.getElementById('tj-usetextjoin').checked?' + TEXTJOIN':'') +
    ' with inlined INDIRECT. No LET. Excel 2019+ (dynamic arrays).';

  const legEl = document.getElementById('tj-legacy-content');
  legEl.innerHTML =
    (r.isFallback ? `<div class="inline-warn">⚠️ Dynamic last-row not available in legacy — range fixed to row ${r.endRow}. Use Manual end row for a precise limit.</div>` : '') +
    `<p class="sub-label">Option A — INDEX + MATCH (first match only)</p>
     <textarea class="formula-textarea leg" id="tj-out-leg-a" readonly>${r.legA}</textarea>
     <div style="display:flex;gap:8px;align-items:center;margin-top:8px;margin-bottom:14px">
       <button class="copy-btn" onclick="copyFormula('tj-out-leg-a',this)">📋 Copy</button>
       <span class="pane-desc" style="margin:0">Returns the first matching value only.</span>
     </div>
     <p class="sub-label">Option B — IF + SMALL + INDEX (multiple matches, array formula)</p>
     <textarea class="formula-textarea leg" id="tj-out-leg-b" readonly>${r.legB}</textarea>
     <div style="display:flex;gap:8px;align-items:center;margin-top:8px">
       <button class="copy-btn" onclick="copyFormula('tj-out-leg-b',this)">📋 Copy</button>
       <span class="pane-desc" style="margin:0">Confirm with <strong style="color:var(--accent4)">Ctrl + Shift + Enter</strong>. Copy down for all matches.</span>
     </div>`;

  document.getElementById('tj-output-section').classList.add('visible');
  tj_switchTab('adv');
  if (mismatchWarn) ['tj-pane-adv','tj-pane-c19'].forEach(p => showPaneWarning(p, mismatchWarn));
  else              ['tj-pane-adv','tj-pane-c19'].forEach(p => clearPaneWarning(p));

  const panel = document.getElementById('tj-explain-panel');
  if (panel?.classList.contains('open')) panel.innerHTML = tj_explainBuildFn();

  document.getElementById('tj-output-section').scrollIntoView({ behavior:'smooth', block:'nearest' });
}
