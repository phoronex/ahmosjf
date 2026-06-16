/* ═══════════════════════════════════════════════════════════
   card-im.js — INDEX + MATCH  (V6.2 — full review pass)
   • Force text coercion: &"" on BOTH lookup value AND lookup array
   • TRIM correctly applied before coercion
   • Live preview  •  Explain  •  Export All
   ═══════════════════════════════════════════════════════════ */
'use strict';

const IM_TABS = ['adv','c19','leg'];
function im_switchTab(tab) { switchTab('im', tab, IM_TABS); }

function im_reset() {
  resetFormByPrefix('im', () => {
    document.getElementById('im-match').selectedIndex = 0;
    document.getElementById('im-addstring').checked   = true;
    document.getElementById('im-lock').checked        = true;
    document.getElementById('im-usetrim').checked     = false;
  });
  const pr = document.getElementById('im-live-preview');
  if (pr) { pr.textContent = '— fill in the fields above to see a live preview —'; pr.className = 'live-preview-box hint'; }
  const panel = document.getElementById('im-explain-panel');
  if (panel) { panel.classList.remove('open'); panel.innerHTML = ''; }
  const btn = document.getElementById('im-explain-btn');
  if (btn) btn.textContent = '💡 Explain Formula';
}

function im_buildAll() {
  const filePath      = document.getElementById('im-filepath').value.trim();
  const sheetName     = document.getElementById('im-sheetname').value.trim();
  const lookupCellRaw = document.getElementById('im-lookupcell').value.trim();
  const returnRaw     = document.getElementById('im-returnarray').value.trim();
  const lookupRaw     = document.getElementById('im-lookuparray').value.trim();
  const matchType     = document.getElementById('im-match').value;
  const ifnaVal       = document.getElementById('im-ifna').value.trim();
  const ierrVal       = document.getElementById('im-error').value.trim();
  const lock          = document.getElementById('im-lock').checked;
  const addStr        = document.getElementById('im-addstring').checked;
  const useTrim       = document.getElementById('im-usetrim').checked;

  if (!lookupCellRaw || !returnRaw || !lookupRaw) return null;

  const lookupCell = normaliseCellRef(lookupCellRaw);
  if (!/^[A-Z]{1,3}\d{1,7}$/.test(lookupCell)) return null;

  const rtV = validateRange(returnRaw, 'Return Array');
  if (!rtV.ok) return null;
  const lkV = validateRange(lookupRaw, 'Lookup Array');
  if (!lkV.ok) return null;

  const sheetPath = buildSheetPath(filePath, sheetName);
  if (sheetPath === null) return null;

  const rtComp = parseRange(returnRaw);
  const lkComp = parseRange(lookupRaw);
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
    const lr = getLastRowInfo('im', lkComp.endCol, sheetPath);
    if (!lr.ok) return null;
    lastRow  = lr.formula;
    needsVar = (lr.mode === 'formula');
  }

  const errText = (ierrVal || 'Not Found');

  // ── Lookup value: TRIM first, then coercion ────────────────
  const lv_trimmed = useTrim ? `TRIM(${lookupCell})` : lookupCell;
  const lv         = addStr ? `${lv_trimmed}&""` : lv_trimmed;

  // ── Lookup array: coerce to text when addStr ON ────────────
  const makeArr = arr => addStr ? `${arr}&""` : arr;

  // ── INDEX+MATCH core ───────────────────────────────────────
  const core = (lkArr, rtArr) =>
    `INDEX(${rtArr},MATCH(${lv},${makeArr(lkArr)},${matchType}))`;

  // ── Build formulas ─────────────────────────────────────────
  const lkI = (lrE) => buildIndirectRange(lkComp, defStart, lrE, sheetPath, lock);
  const rtI = (lrE) => buildIndirectRange(rtComp, defStart, lrE, sheetPath, lock);

  let adv;
  if (needsVar) {
    adv = `=LET(lastRow,${lastRow},` +
          `lkArr,${lkI('lastRow')},` +
          `rtArr,${rtI('lastRow')},` +
          `${wrapErrors(core('lkArr','rtArr'), ifnaVal, errText)})`;
  } else {
    adv = `=${wrapErrors(core(lkI(lastRow), rtI(lastRow)), ifnaVal, errText)}`;
  }

  const c19 = `=${wrapErrors(core(lkI(lastRow), rtI(lastRow)), ifnaVal, errText)}`;

  const endRow  = explicitEnd || '10000';
  const lkStatic = buildStaticRange(lkComp, defStart, endRow, sheetPath, lock);
  const rtStatic = buildStaticRange(rtComp, defStart, endRow, sheetPath, lock);
  const leg = `=${wrapErrors(core(lkStatic, rtStatic), ifnaVal, errText)}`;

  const legWarns = [];
  if (!explicitEnd && needsVar)
    legWarns.push(`⚠️ Dynamic last-row not available in legacy — fixed to row ${endRow}. Use Manual end row for accuracy.`);
  if (mismatchWarn) legWarns.push(mismatchWarn);

  const coercionNote = addStr
    ? '📝 <strong>Force text ON — both sides:</strong> Lookup value coerced with <code>&amp;""</code> AND lookup array coerced with <code>&amp;""</code>. Ensures text-stored numbers and actual numbers both match correctly.'
    : '📝 <strong>Force text OFF:</strong> No coercion. Data types must match between your lookup value and lookup array.';

  return {
    adv, c19, leg, mismatchWarn, legWarns, coercionNote,
    advDesc : `LET + INDIRECT + INDEX + MATCH. Dynamic range, named variables. Excel 365 / 2021+.${addStr ? ' Both-sides coercion ON.' : ''}`,
    c19Desc : `INDIRECT + INDEX + MATCH. Dynamic range. Excel 2007+.${addStr ? ' Both-sides coercion ON.' : ''}`,
    legDesc : `Plain INDEX + MATCH. Static range. Excel 2003+.${addStr ? ' Lookup value coerced with &"".' : ''}`,
  };
}

function im_previewBuildFn() {
  try { const r = im_buildAll(); return r?.adv || null; } catch(e) { return null; }
}

function im_explainBuildFn() {
  const addStr  = document.getElementById('im-addstring')?.checked;
  const useTrim = document.getElementById('im-usetrim')?.checked;
  const r       = im_buildAll();
  const notes   = [
    r?.coercionNote || '',
    useTrim ? '✂️ <strong>TRIM ON:</strong> Whitespace stripped from the lookup value before coercion and search.' : '',
  ].filter(Boolean);
  return buildExplainHTML(
    ['LET','INDIRECT','INDEX','MATCH','IFERROR','IFNA'],
    'Uses MATCH to find the position of a value in a lookup array, then INDEX to return ' +
    'the value at that same position from a separate return array. ' +
    'More flexible than VLOOKUP — can return from any column.',
    notes
  );
}

document.addEventListener('DOMContentLoaded', () => {
  ['im-filepath','im-sheetname','im-lookupcell','im-returnarray',
   'im-lookuparray','im-ifna','im-error','im-manualrow']
    .forEach(id => document.getElementById(id)
      ?.addEventListener('input', () => schedulePreview('im', im_previewBuildFn)));
  ['im-match','im-addstring','im-lock','im-usetrim']
    .forEach(id => document.getElementById(id)
      ?.addEventListener('change', () => schedulePreview('im', im_previewBuildFn)));
  document.querySelectorAll('input[name="im-lastrow"]')
    .forEach(r => r.addEventListener('change',
      () => schedulePreview('im', im_previewBuildFn)));
});

function generateIndexMatch() {
  const lookupCellRaw = document.getElementById('im-lookupcell').value.trim();
  const returnRaw     = document.getElementById('im-returnarray').value.trim();
  const lookupRaw     = document.getElementById('im-lookuparray').value.trim();
  const filePath      = document.getElementById('im-filepath').value.trim();
  const sheetName     = document.getElementById('im-sheetname').value.trim();

  if (!lookupCellRaw || !returnRaw || !lookupRaw) {
    alert('Please fill in: Lookup Cell, Return Array, and Lookup Array.'); return;
  }
  const lookupNorm = normaliseCellRef(lookupCellRaw);
  if (!/^[A-Z]{1,3}\d{1,7}$/.test(lookupNorm)) {
    alert(`Lookup Cell "${lookupCellRaw}" is not a valid cell reference.\nExample: A2`); return;
  }
  const rtV = validateRange(returnRaw, 'Return Array');
  if (!rtV.ok) { alert(rtV.error); return; }
  const lkV = validateRange(lookupRaw, 'Lookup Array');
  if (!lkV.ok) { alert(lkV.error); return; }
  const sp = buildSheetPath(filePath, sheetName);
  if (sp === null) { alert(processFilePath(filePath).error); return; }

  const r = im_buildAll();
  if (!r) { alert('Could not build formula — check your inputs.'); return; }

  document.getElementById('im-out-adv').value        = r.adv;
  document.getElementById('im-desc-adv').textContent = r.advDesc;
  document.getElementById('im-out-c19').value        = r.c19;
  document.getElementById('im-desc-c19').textContent = r.c19Desc;
  document.getElementById('im-out-leg').value        = r.leg;
  document.getElementById('im-desc-leg').textContent = r.legDesc;

  ['adv','c19'].forEach(t =>
    r.mismatchWarn
      ? showPaneWarning(`im-pane-${t}`, r.mismatchWarn)
      : clearPaneWarning(`im-pane-${t}`)
  );
  r.legWarns.length
    ? showPaneWarning('im-pane-leg', r.legWarns.join('\n\n'))
    : clearPaneWarning('im-pane-leg');

  document.getElementById('im-output-section').classList.add('visible');
  im_switchTab('adv');

  const panel = document.getElementById('im-explain-panel');
  if (panel?.classList.contains('open')) panel.innerHTML = im_explainBuildFn();

  document.getElementById('im-output-section')
    .scrollIntoView({ behavior:'smooth', block:'nearest' });
}
