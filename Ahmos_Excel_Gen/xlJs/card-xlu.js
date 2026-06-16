/* ═══════════════════════════════════════════════════════════
   card-xlu.js — XLOOKUP  (V6.2 — full review pass)
   • Force text coercion: &"" on BOTH lookup value AND lookup array
   • TRIM correctly applied before coercion when checked
   • Live preview  •  Explain  •  Export All
   ═══════════════════════════════════════════════════════════ */
'use strict';

const XLU_TABS = ['adv','c19','leg'];
function xlu_switchTab(tab) { switchTab('xlu', tab, XLU_TABS); }

function xlu_reset() {
  resetFormByPrefix('xlu', () => {
    document.getElementById('xlu-matchmode').selectedIndex  = 0;
    document.getElementById('xlu-searchmode').selectedIndex = 0;
    document.getElementById('xlu-addstring').checked = true;
    document.getElementById('xlu-lock').checked      = true;
    document.getElementById('xlu-usetrim').checked   = false;
  });
  const pr = document.getElementById('xlu-live-preview');
  if (pr) { pr.textContent = '— fill in the fields above to see a live preview —'; pr.className = 'live-preview-box hint'; }
  const panel = document.getElementById('xlu-explain-panel');
  if (panel) { panel.classList.remove('open'); panel.innerHTML = ''; }
  const btn = document.getElementById('xlu-explain-btn');
  if (btn) btn.textContent = '💡 Explain Formula';
}

function xlu_buildAll() {
  const filePath     = document.getElementById('xlu-filepath').value.trim();
  const sheetName    = document.getElementById('xlu-sheetname').value.trim();
  const lookupValRaw = document.getElementById('xlu-lookupval').value.trim();
  const lookupColRaw = document.getElementById('xlu-lookupcol').value.trim();
  const returnColRaw = document.getElementById('xlu-returncol').value.trim();
  const ifnaVal      = document.getElementById('xlu-ifna').value.trim();
  const ierrVal      = document.getElementById('xlu-iferror').value.trim();
  const matchMode    = document.getElementById('xlu-matchmode').value;
  const searchMode   = document.getElementById('xlu-searchmode').value;
  const addStr       = document.getElementById('xlu-addstring').checked;
  const lock         = document.getElementById('xlu-lock').checked;
  const useTrim      = document.getElementById('xlu-usetrim').checked;

  if (!lookupValRaw || !lookupColRaw || !returnColRaw) return null;

  const lkV = validateRange(lookupColRaw, 'Lookup Column'); if (!lkV.ok) return null;
  const rtV = validateRange(returnColRaw, 'Return Column'); if (!rtV.ok) return null;

  const lookupMV  = resolveMatchValue(lookupValRaw);
  const sheetPath = buildSheetPath(filePath, sheetName);
  if (sheetPath === null) return null;

  const lkComp = parseRange(lookupColRaw);
  const rtComp = parseRange(returnColRaw);
  const defStart     = String(Math.max(
    parseInt(lkComp.startRow || '1'),
    parseInt(rtComp.startRow || '1')
  ));
  const mismatchWarn = warnIfStartRowMismatch(lkComp, rtComp);

  const explicitEnd = lkComp.endRow || rtComp.endRow;
  let lastRow, needsVar = false;
  if (explicitEnd) {
    lastRow = explicitEnd;
  } else {
    const lr = getLastRowInfo('xlu', lkComp.endCol, sheetPath);
    if (!lr.ok) return null;
    lastRow  = lr.formula;
    needsVar = (lr.mode === 'formula');
  }

  // ── Lookup value: apply TRIM first, then coercion ─────────
  // For a cell ref: TRIM(A2)&""  when both TRIM and addStr are ON
  // For a literal string: already quoted text — coercion irrelevant
  const lv_trimmed = (lookupMV.type === 'cell' && useTrim)
    ? `TRIM(${lookupMV.formula})`
    : lookupMV.formula;

  const lv = (addStr && lookupMV.type === 'cell')
    ? `${lv_trimmed}&""`   // coerce cell-ref lookup to text
    : lv_trimmed;           // literal strings are already text; no coercion needed

  // ── Lookup array: coerce to text when addStr ON ────────────
  const makeLookupArr = arr => addStr ? `${arr}&""` : arr;

  // ── XLOOKUP core — not_found arg handles the "not found" case ─
  const nf = (ifnaVal || '').replace(/"/g, '""');

  const xlCore = (lkArr, rtArr) =>
    `XLOOKUP(${lv},${makeLookupArr(lkArr)},${rtArr},"${nf}",${matchMode},${searchMode})`;

  // ── Build output ranges ────────────────────────────────────
  const lkIndirect  = (lrE) => buildIndirectRange(lkComp, defStart, lrE, sheetPath, lock);
  const rtIndirect  = (lrE) => buildIndirectRange(rtComp, defStart, lrE, sheetPath, lock);

  // Wrap outer IFERROR (catches structural errors; XLOOKUP already handles not-found)
  const wrap = (inner) => ierrVal
    ? `IFERROR(${inner},"${ierrVal.replace(/"/g,'""')}")`
    : inner;

  let adv, c19;
  if (needsVar) {
    adv = `=LET(lastRow,${lastRow},` +
          `lkArr,${lkIndirect('lastRow')},` +
          `rtArr,${rtIndirect('lastRow')},` +
          `${wrap(xlCore('lkArr','rtArr'))})`;
  } else {
    adv = `=${wrap(xlCore(lkIndirect(lastRow), rtIndirect(lastRow)))}`;
  }
  c19 = `=${wrap(xlCore(lkIndirect(lastRow), rtIndirect(lastRow)))}`;

  // ── Legacy: INDEX + MATCH fallback ────────────────────────
  const endRow = explicitEnd || '10000';
  const lkStatic = buildStaticRange(lkComp, defStart, endRow, sheetPath, lock);
  const rtStatic = buildStaticRange(rtComp, defStart, endRow, sheetPath, lock);
  const errStr   = (ierrVal || ifnaVal || 'Not Found').replace(/"/g, '""');
  const leg = `=IFERROR(INDEX(${rtStatic},MATCH(${lv},${makeLookupArr(lkStatic)},0)),"${errStr}")`;

  const legWarns = [];
  if (!explicitEnd && needsVar)
    legWarns.push(`⚠️ Dynamic last-row not available in legacy — range fixed to row ${endRow}. Use Manual end row for accuracy.`);
  if (matchMode === '2')
    legWarns.push('⚠️ Wildcard match (mode 2) is not supported by INDEX+MATCH. Exact match (0) used instead.');
  if (searchMode === '-1')
    legWarns.push('⚠️ Reverse search (Last-to-First) is not supported by INDEX+MATCH. First-to-Last used instead.');
  if (mismatchWarn) legWarns.push(mismatchWarn);

  const coercionNote = addStr
    ? '📝 <strong>Force text ON — both sides:</strong> Lookup value coerced with <code>&amp;""</code> (cell refs only) AND lookup array coerced with <code>&amp;""</code>. Ensures text and numeric versions of the same value match correctly.'
    : '📝 <strong>Force text OFF:</strong> No coercion. Ensure your lookup value and lookup array share the same data type to avoid #N/A.';

  return {
    adv, c19, leg, mismatchWarn, legWarns, coercionNote,
    advDesc : `LET + INDIRECT + XLOOKUP. Dynamic range, named variables. Excel 365 / 2021+.${addStr ? ' Both-sides coercion ON.' : ''}`,
    c19Desc : `XLOOKUP + INDIRECT (no LET). Excel 365 / 2021+.${addStr ? ' Both-sides coercion ON.' : ''}`,
    legDesc : 'INDEX + MATCH fallback (no XLOOKUP). Exact match only. Excel 2007+.',
  };
}

function xlu_previewBuildFn() {
  try { const r = xlu_buildAll(); return r?.adv || null; } catch(e) { return null; }
}

function xlu_explainBuildFn() {
  const addStr  = document.getElementById('xlu-addstring')?.checked;
  const useTrim = document.getElementById('xlu-usetrim')?.checked;
  const r       = xlu_buildAll();
  const notes   = [
    r?.coercionNote || '',
    useTrim ? '✂️ <strong>TRIM ON:</strong> Whitespace stripped from the lookup value before coercion and search.' : '',
  ].filter(Boolean);
  return buildExplainHTML(
    ['LET','INDIRECT','XLOOKUP','IFERROR'],
    'Finds a value in a lookup column and returns a value from a separate return column. ' +
    'Unlike VLOOKUP, it can look left, search in reverse, and support wildcards.',
    notes
  );
}

document.addEventListener('DOMContentLoaded', () => {
  ['xlu-filepath','xlu-sheetname','xlu-lookupval','xlu-lookupcol',
   'xlu-returncol','xlu-ifna','xlu-iferror','xlu-manualrow']
    .forEach(id => document.getElementById(id)
      ?.addEventListener('input', () => schedulePreview('xlu', xlu_previewBuildFn)));
  ['xlu-matchmode','xlu-searchmode','xlu-addstring','xlu-lock','xlu-usetrim']
    .forEach(id => document.getElementById(id)
      ?.addEventListener('change', () => schedulePreview('xlu', xlu_previewBuildFn)));
  document.querySelectorAll('input[name="xlu-lastrow"]')
    .forEach(r => r.addEventListener('change',
      () => schedulePreview('xlu', xlu_previewBuildFn)));
});

function generateXlookup() {
  const filePath     = document.getElementById('xlu-filepath').value.trim();
  const lookupValRaw = document.getElementById('xlu-lookupval').value.trim();
  const lookupColRaw = document.getElementById('xlu-lookupcol').value.trim();
  const returnColRaw = document.getElementById('xlu-returncol').value.trim();

  if (!lookupValRaw || !lookupColRaw || !returnColRaw) {
    alert('Please fill in: Lookup Value, Lookup Column, and Return Column.'); return;
  }
  const lkV = validateRange(lookupColRaw, 'Lookup Column');
  if (!lkV.ok) { alert(lkV.error); return; }
  const rtV = validateRange(returnColRaw, 'Return Column');
  if (!rtV.ok) { alert(rtV.error); return; }
  const sp = buildSheetPath(filePath, document.getElementById('xlu-sheetname').value.trim());
  if (sp === null) { alert(processFilePath(filePath).error); return; }

  const r = xlu_buildAll();
  if (!r) { alert('Could not build formula — check your inputs.'); return; }

  document.getElementById('xlu-out-adv').value        = r.adv;
  document.getElementById('xlu-desc-adv').textContent = r.advDesc;
  document.getElementById('xlu-out-c19').value        = r.c19;
  document.getElementById('xlu-desc-c19').textContent = r.c19Desc;
  document.getElementById('xlu-out-leg').value        = r.leg;
  document.getElementById('xlu-desc-leg').textContent = r.legDesc;

  ['adv','c19'].forEach(t =>
    r.mismatchWarn
      ? showPaneWarning(`xlu-pane-${t}`, r.mismatchWarn)
      : clearPaneWarning(`xlu-pane-${t}`)
  );
  r.legWarns.length
    ? showPaneWarning('xlu-pane-leg', r.legWarns.join('\n\n'))
    : clearPaneWarning('xlu-pane-leg');

  document.getElementById('xlu-output-section').classList.add('visible');
  xlu_switchTab('adv');

  const panel = document.getElementById('xlu-explain-panel');
  if (panel?.classList.contains('open')) panel.innerHTML = xlu_explainBuildFn();

  document.getElementById('xlu-output-section')
    .scrollIntoView({ behavior:'smooth', block:'nearest' });
}
